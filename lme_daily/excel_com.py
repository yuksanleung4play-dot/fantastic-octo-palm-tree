"""Excel COM 共用封裝（Windows + 本機 Excel）。

永遠只沿用已開啟的 Excel（GetActiveObject）。禁止 Dispatch / DispatchEx
另開行程：那是 Bloomberg Terminal 被鎖的根因。結束時不 Quit。
"""

from __future__ import annotations

import logging
import os
import sys
import time
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import Any

from lme_daily.exceptions import ExcelComError
from lme_daily.unc_paths import rewrite_p_drive_to_unc, rewrite_unc_to_p_drive

logger = logging.getLogger(__name__)

XL_DONE = 0
XL_CALCULATING = 1
XL_PENDING = 2

# RPC_E_DISCONNECTED — 工作簿已被巨集 / 使用者關閉
RPC_E_DISCONNECTED = -2147417848  # 0x80010108

_workbook_open_count = 0


def reset_workbook_open_count() -> None:
    global _workbook_open_count
    _workbook_open_count = 0


def get_workbook_open_count() -> int:
    """本次行程實際呼叫 ``Workbooks.Open`` 的次數（沿用已開啟檔不計）。"""
    return _workbook_open_count


def _record_workbook_open() -> None:
    global _workbook_open_count
    _workbook_open_count += 1


def require_windows_excel() -> None:
    if sys.platform != "win32":
        raise ExcelComError(
            f"目前作業系統是 {sys.platform}。開啟 Excel / 呼叫 VBA / 刷新 Bloomberg "
            "必須在已安裝 Excel 與 Bloomberg Terminal 的 Windows 上執行。"
        )


def import_win32com():
    require_windows_excel()
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise ExcelComError(
            "找不到 pywin32。請在 Windows 上執行：pip install pywin32"
        ) from exc
    return win32com.client, pythoncom


@contextmanager
def excel_app(
    *,
    visible: bool = True,
    display_alerts: bool = False,
    reuse_running: bool = True,
    quit_on_exit: bool = False,
    new_instance: bool = False,
) -> Iterator[Any]:
    """取得既有 Excel.Application（GetActiveObject）。

    找不到已開啟的 Excel 就報錯，**絕不** Dispatch / DispatchEx 開新進程。
    ``yield`` 期間呼叫端拋出的例外維持原樣，不會被包成「啟動 Excel 失敗」。
    """
    if not reuse_running:
        logger.warning("excel.reuse_running=false 已忽略：禁止另開 Excel 進程")
    if quit_on_exit:
        logger.warning("excel.quit_on_exit=true 已忽略：不會 Quit 沿用的 Excel")

    win32com_client, pythoncom = import_win32com()
    pythoncom.CoInitialize()
    app = None
    try:
        try:
            app = _acquire_running_excel(win32com_client, new_instance=new_instance)
        except ExcelComError:
            raise
        except Exception as exc:
            raise ExcelComError(f"啟動 Excel 失敗：{exc}") from exc
        _log_attached_excel_process(app)
        _probe_p_drive_in_excel(app)
        app.Visible = visible
        app.DisplayAlerts = display_alerts
        try:
            app.AskToUpdateLinks = False
        except Exception:
            logger.debug("AskToUpdateLinks 無法設定，略過")
        yield app
    finally:
        if app is not None:
            logger.info("Excel 保持開啟，不呼叫 Quit（避免 Bloomberg Terminal 被鎖）")
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _acquire_running_excel(win32com_client: Any, *, new_instance: bool) -> Any:
    """只允許 GetActiveObject。DispatchEx / Dispatch 分支已刪除，避免以後改回去。"""
    if new_instance:
        raise ExcelComError(
            "excel.new_instance 已停用：禁止 DispatchEx 另開 Excel 進程"
            "（會導致 Bloomberg Terminal 上鎖）。請先手動開 Excel，"
            "登入並解鎖 Bloomberg Terminal 後再執行。"
        )
    try:
        app = win32com_client.GetActiveObject("Excel.Application")
    except Exception as exc:
        raise ExcelComError(
            "找不到已開啟的 Excel。請先手動開 Excel，登入並解鎖 Bloomberg Terminal，"
            "並手動打開「LME BBG WORKBOOK.xlsx」後再執行。"
            f" 原始錯誤：{exc}"
        ) from exc
    logger.info("沿用已開啟的 Excel.Application（GetActiveObject）；沒有 Dispatch/DispatchEx fallback")
    return app


def _log_attached_excel_process(app: Any) -> None:
    """確認腳本連到的是使用者手動開著的那個 Excel 進程。"""
    caption = "<unknown>"
    hwnd: int | None = None
    pid: int | None = None
    try:
        caption = str(app.Caption)
    except Exception as exc:
        logger.debug("讀取 Excel Caption 失敗：%s", exc)
    try:
        hwnd = int(app.Hwnd)
    except Exception as exc:
        logger.debug("讀取 Excel Hwnd 失敗：%s", exc)
    if hwnd is not None:
        try:
            import win32process  # type: ignore

            _, pid = win32process.GetWindowThreadProcessId(hwnd)
        except Exception as exc:
            logger.debug("GetWindowThreadProcessId 失敗：%s", exc)
    logger.info(
        "Excel 進程 PID=%s, Hwnd=%s, Caption=%s；Python PID=%s",
        pid,
        hwnd,
        caption,
        os.getpid(),
    )


def _probe_p_drive_in_excel(app: Any) -> None:
    """Excel 4.0 FILES 只是輔助；現代 Excel 常關掉 Excel 4.0 巨集，失敗不代表沒有 P:。

    準確測試請看 ``TestPDriveVisible``（Dir）。GetActiveObject 沿用的就是你 Alt+F8
    那個 Excel，P: 通常看得到；不要因為 FILES 失敗就把 QueryTable 改成 UNC。
    """
    try:
        result = app.ExecuteExcel4Macro('FILES("P:\\*")')
    except Exception as exc:
        logger.info(
            "Excel 4.0 FILES 探測失敗（此 API 常被安全性關掉，不代表 P: 真的看不到）：%s",
            exc,
        )
        return
    if not isinstance(result, (str, bool, int, float, type(None))):
        logger.debug("P: FILES 探測回傳非純量（%s），略過解讀", type(result).__name__)
        return
    if result in (False, None, "", 0):
        logger.info(
            "Excel 4.0 FILES 回傳 %r（不代表 P: 看不到；請看 TestPDriveVisible）",
            result,
        )
        return
    logger.info("Excel 4.0 FILES 看得到 P:，樣本：%s", result)


P_DRIVE_TEST_MACRO = "TestPDriveVisible"


def run_p_drive_visibility_macro(app: Any, workbook: Any) -> str | None:
    """執行參考工作簿的 TestPDriveVisible；沒有該函式則略過。"""
    wb_name = str(getattr(workbook, "Name", "") or "")
    names: list[str] = []
    if wb_name:
        names.append(f"'{wb_name}'!{P_DRIVE_TEST_MACRO}")
    names.append(P_DRIVE_TEST_MACRO)
    last_exc: BaseException | None = None
    for name in names:
        try:
            result = app.Run(name)
            logger.info("P: 磁碟機可見性測試結果：%s", result)
            return str(result)
        except Exception as exc:
            last_exc = exc
            logger.debug("Run(%s) 失敗：%s", name, exc)
    logger.warning(
        "無法執行 TestPDriveVisible（請把 examples/TestPDriveVisible.bas 匯入參考工作簿）。"
        "最後錯誤：%s",
        last_exc,
    )
    return None


def _iter_com_collection(coll: Any) -> Iterator[Any]:
    try:
        count = int(coll.Count)
    except Exception:
        return
    for index in range(1, count + 1):
        try:
            yield coll.Item(index)
        except Exception:
            try:
                yield coll(index)
            except Exception:
                continue


_QT_PATH_ATTRS = ("Connection", "CommandText", "TextFileName")


def _rewrite_string_attr(obj: Any, attr: str, transform, *, from_label: str, to_label: str) -> bool:
    try:
        current = getattr(obj, attr)
    except Exception:
        return False
    if not isinstance(current, str) or not current:
        return False
    updated = transform(current)
    if updated == current:
        return False
    try:
        setattr(obj, attr, updated)
    except Exception as exc:
        logger.warning("無法改寫 %s.%s（%s）", type(obj).__name__, attr, exc)
        return False
    logger.info("已把 %s 從 %s 改成 %s：%s", attr, from_label, to_label, updated)
    return True


def _iter_query_targets(workbook: Any) -> Iterator[Any]:
    try:
        sheets = workbook.Worksheets
    except Exception as exc:
        logger.debug("無法列舉 Worksheets：%s", exc)
        sheets = None
    if sheets is not None:
        for ws in _iter_com_collection(sheets):
            for coll_name in ("QueryTables", "ListObjects"):
                try:
                    coll = getattr(ws, coll_name)
                except Exception:
                    continue
                for item in _iter_com_collection(coll):
                    target = item
                    if coll_name == "ListObjects":
                        try:
                            target = item.QueryTable
                        except Exception:
                            continue
                    yield target
    try:
        connections = workbook.Connections
    except Exception:
        return
    for conn in _iter_com_collection(connections):
        for sub_name in ("OLEDBConnection", "ODBCConnection", "TextConnection"):
            try:
                yield getattr(conn, sub_name)
            except Exception:
                continue


def log_workbook_query_paths(workbook: Any) -> list[str]:
    """跑巨集前把 QueryTable / 連線路徑打到 log，方便對照 1004 找不到 dat。"""
    logged: list[str] = []
    for target in _iter_query_targets(workbook):
        for attr in _QT_PATH_ATTRS:
            try:
                value = getattr(target, attr)
            except Exception:
                continue
            if isinstance(value, str) and value:
                line = f"{type(target).__name__}.{attr}={value}"
                logged.append(line)
                logger.info("Query 路徑 %s", line)
    if not logged:
        logger.info("工作簿目前沒有可讀的 QueryTable Connection / TextFileName")
    return logged


def apply_path_transform_to_workbook(
    workbook: Any,
    transform,
    *,
    from_label: str,
    to_label: str,
) -> int:
    """記憶體改寫 QueryTable / 連線 / VBA，不存檔。"""
    changed = 0
    for target in _iter_query_targets(workbook):
        for attr in _QT_PATH_ATTRS:
            if _rewrite_string_attr(
                target, attr, transform, from_label=from_label, to_label=to_label
            ):
                changed += 1
    changed += _rewrite_vba_paths(
        workbook, transform, from_label=from_label, to_label=to_label
    )
    if changed:
        logger.info(
            "已把參考工作簿內 %d 處路徑從 %s 改成 %s（未存檔）",
            changed,
            from_label,
            to_label,
        )
    else:
        logger.info(
            "參考工作簿未發現可從 %s 改成 %s 的路徑（或 VBProject 無法存取）",
            from_label,
            to_label,
        )
    return changed


def rewrite_workbook_p_drive_to_unc(workbook: Any) -> int:
    """可選：把 P: 改成 UNC。TEXT QueryTable 對 UNC 常 1004，預設不要用。"""
    logger.warning(
        "正在把 QueryTable/VBA 的 P: 改成 UNC。Excel TEXT 連線對 UNC 常回報找不到 .dat；"
        "若巨集 1004，請把 vba.rewrite_p_drive_to_unc 改回 false。"
    )
    return apply_path_transform_to_workbook(
        workbook,
        rewrite_p_drive_to_unc,
        from_label="P:",
        to_label="UNC",
    )


def rewrite_workbook_unc_to_p_drive(workbook: Any) -> int:
    """把上一版改成 UNC 的 TEXT 連線改回 P:，讓 QueryTables.Refresh 找得到 .dat。"""
    return apply_path_transform_to_workbook(
        workbook,
        rewrite_unc_to_p_drive,
        from_label="UNC",
        to_label="P:",
    )


def _rewrite_vba_paths(workbook: Any, transform, *, from_label: str, to_label: str) -> int:
    try:
        vbproj = workbook.VBProject
        components = vbproj.VBComponents
    except Exception as exc:
        logger.warning(
            "無法讀取 VBProject 以改寫路徑（%s → %s）。請在 Excel 啟用"
            "「信任存取 VBA 專案物件模型」。錯誤：%s",
            from_label,
            to_label,
            exc,
        )
        return 0
    changed = 0
    for comp in _iter_com_collection(components):
        try:
            module = comp.CodeModule
            n_lines = int(module.CountOfLines)
        except Exception:
            continue
        if n_lines <= 0:
            continue
        try:
            original = str(module.Lines(1, n_lines))
        except Exception:
            continue
        updated = transform(original)
        if updated == original:
            continue
        try:
            module.DeleteLines(1, n_lines)
            module.AddFromString(updated)
        except Exception as exc:
            logger.warning("無法改寫 VBA 模組 %s：%s", getattr(comp, "Name", "?"), exc)
            continue
        changed += 1
        logger.info(
            "已把 VBA 模組 %s 的路徑從 %s 改成 %s",
            getattr(comp, "Name", "?"),
            from_label,
            to_label,
        )
    return changed


def find_open_workbook(app: Any, path: Path) -> Any | None:
    """若目標檔已在此 Excel 開啟則回傳該 workbook。"""
    target = str(path.resolve()).lower()
    name = path.name.lower()
    try:
        for workbook in app.Workbooks:
            try:
                full = str(workbook.FullName).lower()
            except Exception:
                full = ""
            wb_name = str(getattr(workbook, "Name", "")).lower()
            if full == target or wb_name == name:
                return workbook
    except Exception as exc:
        logger.debug("列舉 Workbooks 失敗：%s", exc)
    return None


def open_workbook(
    app: Any,
    path: Path,
    *,
    read_only: bool = False,
    update_links: int = 0,
) -> tuple[Any, bool]:
    """開啟工作簿。回傳 ``(workbook, opened_by_us)``。"""
    resolved = path.resolve()
    existing = find_open_workbook(app, resolved)
    if existing is not None:
        logger.info("工作簿已在 Excel 開啟，沿用：%s", getattr(existing, "FullName", resolved))
        return existing, False
    try:
        present = resolved.is_file()
    except OSError as exc:
        logger.debug("is_file(%s) 失敗（可能是網路磁碟鎖定）：%s", resolved, exc)
        present = True
    if not present:
        raise ExcelComError(f"工作簿不存在：{resolved}")
    logger.info("開啟工作簿：%s", resolved)
    _record_workbook_open()
    try:
        workbook = app.Workbooks.Open(
            str(resolved),
            UpdateLinks=update_links,
            ReadOnly=read_only,
        )
        return workbook, True
    except Exception as exc:
        raise ExcelComError(f"開啟工作簿失敗（{resolved}）：{exc}") from exc


def com_hresult_codes(exc: BaseException) -> set[int]:
    codes: set[int] = set()
    args = getattr(exc, "args", ())
    if args:
        try:
            codes.add(int(args[0]))
        except (TypeError, ValueError):
            pass
    if len(args) >= 3 and isinstance(args[2], (tuple, list)) and len(args[2]) >= 6:
        try:
            if args[2][5] is not None:
                codes.add(int(args[2][5]))
        except (TypeError, ValueError):
            pass
    hresult = getattr(exc, "hresult", None)
    if isinstance(hresult, int):
        codes.add(hresult)
    return codes


def is_rpc_disconnected(exc: BaseException | None) -> bool:
    current: BaseException | None = exc
    seen: set[int] = set()
    for _ in range(6):
        if current is None:
            break
        ident = id(current)
        if ident in seen:
            break
        seen.add(ident)
        if RPC_E_DISCONNECTED in com_hresult_codes(current):
            return True
        current = current.__cause__ or current.__context__
    return False


def _workbook_label(workbook: Any) -> str:
    try:
        name = getattr(workbook, "Name", None)
        if name:
            return str(name)
    except Exception as exc:
        if is_rpc_disconnected(exc):
            return "<disconnected>"
        logger.debug("讀取 workbook.Name 失敗：%s", exc)
        return "<unknown>"
    return "<unknown>"


def close_workbook(workbook: Any, *, save_changes: bool = False) -> None:
    """關閉工作簿。若已被外部關閉（RPC_E_DISCONNECTED）只記 log，不拋例外。"""
    name = _workbook_label(workbook)
    try:
        workbook.Close(SaveChanges=save_changes)
        logger.info("已關閉工作簿：%s（SaveChanges=%s）", name, save_changes)
    except Exception as exc:
        if is_rpc_disconnected(exc):
            logger.info("工作簿已由外部關閉，略過 Close（%s）：%s", name, exc)
            return
        logger.warning("關閉工作簿失敗，略過（%s）：%s", name, exc)


def close_workbook_if_opened(workbook: Any, opened_by_us: bool, *, save_changes: bool = False) -> None:
    if opened_by_us:
        close_workbook(workbook, save_changes=save_changes)
        return
    name = _workbook_label(workbook)
    logger.info("工作簿原本就開著，不關閉：%s", name)


def wait_until_calculation_done(
    app: Any,
    *,
    timeout_seconds: float,
    poll_interval: float = 0.5,
) -> None:
    deadline = time.monotonic() + timeout_seconds
    last_state = None
    while time.monotonic() < deadline:
        try:
            last_state = int(app.CalculationState)
        except Exception as exc:
            raise ExcelComError(f"讀取 CalculationState 失敗：{exc}") from exc
        if last_state == XL_DONE:
            logger.info("Excel CalculationState=xlDone")
            return
        time.sleep(poll_interval)
    raise ExcelComError(
        f"等待 Excel 計算完成逾時（{timeout_seconds:.0f}s），最後狀態={last_state} "
        f"（0=xlDone, 1=xlCalculating, 2=xlPending）"
    )


def wait_for_file(path: Path, *, timeout_seconds: float, poll_interval: float) -> Path:
    return wait_for_any_file([path], timeout_seconds=timeout_seconds, poll_interval=poll_interval)


def wait_for_any_file(
    paths: list[Path],
    *,
    timeout_seconds: float,
    poll_interval: float,
) -> Path:
    unique: list[Path] = []
    for path in paths:
        if path not in unique:
            unique.append(path)
    if not unique:
        raise ExcelComError("wait_for_any_file 沒有任何路徑")
    deadline = time.monotonic() + timeout_seconds
    shown = " 或 ".join(str(p) for p in unique)
    logger.info("等待輸出檔產生：%s（timeout=%.0fs）", shown, timeout_seconds)
    while time.monotonic() < deadline:
        for path in unique:
            try:
                if path.exists() and path.stat().st_size > 0 and _file_unlocked(path):
                    logger.info("輸出檔已就緒：%s（%d bytes）", path, path.stat().st_size)
                    return path
            except OSError:
                logger.debug("檢查檔案時 OSError，繼續等待：%s", path)
        time.sleep(poll_interval)
    raise ExcelComError(
        f"等待檔案逾時（{timeout_seconds:.0f}s）仍未產生或無法讀取：{shown}"
    )


def _file_unlocked(path: Path) -> bool:
    try:
        with path.open("rb"):
            return True
    except OSError:
        return False
