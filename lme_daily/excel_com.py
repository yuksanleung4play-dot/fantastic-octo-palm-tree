"""Excel COM 共用封裝（Windows + 本機 Excel）。

永遠只沿用已開啟的 Excel（GetActiveObject）。禁止 Dispatch / DispatchEx
另開行程：那是 Bloomberg Terminal 被鎖的根因。結束時不 Quit。
"""

from __future__ import annotations

import logging
import sys
import time
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import Any

from lme_daily.exceptions import ExcelComError

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
    logger.info("沿用已開啟的 Excel.Application（GetActiveObject）")
    return app


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
