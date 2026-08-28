"""開啟早班 LME reference 工作簿並執行 VBA 巨集。

``Application.Run(macro, 上日, 3M)`` 若因參數數量不符失敗（DISP_E_BADPARAMCOUNT，
例如 Sub lme_main() 沒有參數、只靠 InputBox），會自動改成：

- 在 COM 執行緒呼叫 ``Run(macro)``（不帶參數）
- 背景執行緒把兩個 InputBox 填成上日日期 / 3M date 並按 Enter

不必再手動把 ``use_param_injection`` 改成 false（仍可用該開關跳過參數嘗試）。
"""

from __future__ import annotations

import logging
import shutil
import threading
import time
from datetime import date, datetime
from pathlib import Path

from lme_daily.config import AppConfig
from lme_daily.dates import calc_lme_dates
from lme_daily.excel_com import (
    excel_app,
    is_rpc_disconnected,
    log_workbook_query_paths,
    open_workbook,
    rewrite_workbook_p_drive_to_unc,
    rewrite_workbook_unc_to_p_drive,
    run_p_drive_visibility_macro,
    wait_for_any_file,
)
from lme_daily.exceptions import ExcelComError, MacroOutputError
from lme_daily.unc_paths import (
    span_dat_candidates,
    working_dir_has_padded_numbered_folders,
)

logger = logging.getLogger(__name__)

# COM HRESULT
DISP_E_EXCEPTION = -2147352567  # 0x80020009
DISP_E_BADPARAMCOUNT = -2147352562  # 0x8002000E  參數個數不對
DISP_E_UNKNOWNNAME = -2147352570  # 0x80020006  找不到名稱
DISP_E_MEMBERNOTFOUND = -2147352573  # 0x80020003


def com_error_codes(exc: BaseException) -> set[int]:
    """抽出 pywintypes.com_error 的 HRESULT（外層 + excepinfo[5]）。"""
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


def is_bad_param_count(exc: BaseException) -> bool:
    """巨集不接受 Application.Run 傳入的額外參數。"""
    current: BaseException | None = exc
    seen: set[int] = set()
    for _ in range(6):
        if current is None:
            break
        ident = id(current)
        if ident in seen:
            break
        seen.add(ident)
        if DISP_E_BADPARAMCOUNT in com_error_codes(current):
            return True
        current = current.__cause__ or current.__context__
    return False


def step2_search_paths(config: AppConfig, as_of: date) -> list[Path]:
    """優先 ``vba_dir\\yyyymmdd.xlsx``，其次 working_dir 根目錄（巨集常寫 ThisWorkbook.Path）。"""
    primary = config.step2_workbook_path(as_of)
    legacy = config.step2_legacy_path(as_of)
    paths = [primary]
    try:
        different = legacy.resolve() != primary.resolve()
    except OSError:
        different = str(legacy) != str(primary)
    if different:
        paths.append(legacy)
    return paths


def relocate_step2_workbook(found: Path, dest: Path) -> Path:
    """若巨集寫到 working_dir 根目錄，搬到 ``vba_dir``。"""
    dest = Path(dest)
    found = Path(found)
    dest.parent.mkdir(parents=True, exist_ok=True)
    try:
        same = found.resolve() == dest.resolve()
    except OSError:
        same = str(found) == str(dest)
    if same:
        return dest
    if dest.exists():
        try:
            dest.unlink()
        except OSError:
            logger.debug("無法刪除既有 dest，稍後改以複製覆寫：%s", dest)
    try:
        shutil.move(str(found), str(dest))
        logger.info("已將 VBA 產出從 %s 移到 %s", found, dest)
        return dest
    except OSError as exc:
        logger.warning("無法搬移 %s → %s（%s），改為複製", found, dest, exc)
        shutil.copy2(found, dest)
        logger.info("已將 VBA 產出從 %s 複製到 %s", found, dest)
        return dest


def preflight_span_dat(config: AppConfig, *, as_of: date, prev_ymd: str) -> Path | None:
    """確認 LME SPAN ``lme.yyyymmdd.dat`` 在 Python 端看得到，避免 Excel 1004 才發現。

    巨集 QueryTables.Refresh 用上日（InputBox yyyymmdd）組檔名；也檢查執行日檔名。
    目錄不存在只警告（資料夾名可能不同）；目錄在但檔都不在則中斷。
    """
    span_dir = config.span_dat_dir()
    candidates = span_dat_candidates(span_dir, (prev_ymd, as_of.strftime("%Y%m%d")))
    logger.info("LME SPAN 目錄：%s", span_dir)
    try:
        dir_ok = span_dir.is_dir()
    except OSError as exc:
        logger.warning("無法存取 LME SPAN 目錄 %s：%s", span_dir, exc)
        return None
    if not dir_ok:
        logger.warning(
            "找不到 LME SPAN 目錄 %s。巨集 Refresh 仍會去找 lme.yyyymmdd.dat。"
            "請確認路徑，或在 config.yaml 設 paths.span_dat_dir。",
            span_dir,
        )
        return None

    found: list[Path] = []
    for path in candidates:
        try:
            exists = path.is_file() and path.stat().st_size > 0
        except OSError:
            exists = False
        if exists:
            found.append(path)
            logger.info("找到 SPAN dat：%s（%d bytes）", path, path.stat().st_size)
        else:
            logger.warning("SPAN dat 不存在：%s", path)

    if not found:
        listed = " ； ".join(str(p) for p in candidates)
        raise ExcelComError(
            "VBA QueryTables.Refresh 會讀 LME SPAN 的 lme.yyyymmdd.dat，但檔案不在。"
            f"已檢查：{listed}。"
            "請把當日/上日 dat 放到該資料夾，或設定 paths.span_dat_dir。"
        )

    primary = span_dir / f"lme.{prev_ymd}.dat"
    if primary not in found:
        logger.warning(
            "上日 %s 的 dat 不在，但找到了其他日期。巨集若用上日檔名仍會 1004。",
            prev_ymd,
        )
    else:
        logger.info(
            "SPAN dat 在 Python 端看得到。若 Excel 仍說找不到，是 QueryTable 路徑格式"
            "（TEXT;UNC 常 1004；請維持 TEXT;P:\\...），不是檔案真的不在。"
        )
    return found[0]


def resolve_existing_step2(config: AppConfig, as_of: date) -> Path | None:
    for path in step2_search_paths(config, as_of):
        try:
            if path.is_file() and path.stat().st_size > 0:
                return path
        except OSError:
            continue
    return None


def _reformat_date_string(value: str, src_fmt: str, dst_fmt: str) -> str | None:
    try:
        return datetime.strptime(value, src_fmt).strftime(dst_fmt)
    except (TypeError, ValueError):
        return None


def run_reference_macro(
    config: AppConfig,
    *,
    as_of: date,
    prev_date: str,
    three_m_date: str,
) -> Path:
    """執行巨集、等到 ``vba_dir\\yyyymmdd.xlsx`` 產生。

    巨集（lme_main）執行完會自己關閉 reference 工作簿。預設
    ``vba.auto_closes_workbook=true``，Python **不要**再 Close，否則會碰到
    RPC_E_DISCONNECTED。成功與否只看中繼檔是否出現且大小正常。
    """
    expected = config.step2_workbook_path(as_of)
    expected.parent.mkdir(parents=True, exist_ok=True)
    if expected.exists():
        logger.warning("輸出檔已存在，將覆寫：%s", expected)

    ibox_prev, ibox_three_python = calc_lme_dates(
        as_of,
        holiday_list=config.holidays,
        date_format=config.vba.inputbox_date_format,
    )
    ibox_three = _reformat_date_string(
        three_m_date, config.vba.date_format, config.vba.inputbox_date_format
    )
    if ibox_three is None:
        ibox_three = ibox_three_python
        logger.warning(
            "無法把注入的 3M date %s 轉成 InputBox 格式，改用 Python 計算值 %s",
            three_m_date,
            ibox_three,
        )
    logger.info(
        "VBA 模式：%s；巨集=%s；參數注入上日/3M=%s / %s；InputBox 上日/3M=%s / %s",
        "先嘗試參數注入，失敗則改填 InputBox"
        if config.vba.use_param_injection
        else "InputBox 自動填入",
        config.vba.macro_name,
        prev_date,
        three_m_date,
        ibox_prev,
        ibox_three,
    )
    logger.info("VBA 中繼檔預期位置（vba_dir）：%s", expected)
    logger.info("working_dir repr=%r", str(config.paths.working_dir))
    if working_dir_has_padded_numbered_folders(config.paths.working_dir):
        logger.warning(
            "working_dir 在「數字.」後面有多個空白（例如 1.      交易部）。"
            "若實際資料夾是「1. 交易部…」只有一個空白，LME SPAN 會找錯目錄。"
            "請對照檔案總管路徑，不要為了對齊註解而補空白。"
        )

    ready = expected
    with excel_app(
        visible=config.excel.visible,
        display_alerts=config.excel.display_alerts,
        reuse_running=config.excel.reuse_running,
        quit_on_exit=config.excel.quit_on_exit,
        new_instance=config.excel.new_instance,
    ) as app:
        workbook, opened_by_us = open_workbook(app, config.paths.ref_workbook)
        try:
            workbook.Activate()
        except Exception:
            logger.debug("Workbook.Activate 失敗，繼續")
        run_p_drive_visibility_macro(app, workbook)
        log_workbook_query_paths(workbook)
        preflight_span_dat(config, as_of=as_of, prev_ymd=ibox_prev)
        if config.vba.rewrite_p_drive_to_unc:
            rewrite_workbook_p_drive_to_unc(workbook)
        else:
            # 上一版把 TEXT;P:\... 改成 UNC，Excel QueryTables.Refresh 會 1004 找不到 dat。
            # 預設改回 P:（記憶體、不存檔）。巨集若執行期再組 P: 路徑，維持原樣即可。
            rewrite_workbook_unc_to_p_drive(workbook)
        log_workbook_query_paths(workbook)
        try:
            _execute_macro(
                app,
                workbook,
                config,
                prev_date=prev_date,
                three_m_date=three_m_date,
                inputbox_prev=ibox_prev,
                inputbox_three=ibox_three,
            )
        except Exception as exc:
            if is_rpc_disconnected(exc) or is_rpc_disconnected(exc.__cause__):
                logger.warning(
                    "巨集執行後參考工作簿已斷線（RPC_E_DISCONNECTED；多半由巨集自行關閉），"
                    "改以輸出檔判定成功：%s",
                    exc,
                )
            elif isinstance(exc, (ExcelComError, MacroOutputError)):
                raise
            else:
                raise ExcelComError(f"執行巨集 {config.vba.macro_name!r} 失敗：{exc}") from exc
        try:
            found = wait_for_any_file(
                step2_search_paths(config, as_of),
                timeout_seconds=config.vba.output_timeout_seconds,
                poll_interval=config.vba.poll_interval_seconds,
            )
            ready = relocate_step2_workbook(found, expected)
        except ExcelComError as exc:
            searched = " ； ".join(str(p) for p in step2_search_paths(config, as_of))
            raise MacroOutputError(
                f"巨集執行後未產生預期檔案 {expected.name}。"
                f"VBA 中繼檔必須落在 {expected.parent}（vba_dir），"
                f"或仍寫在 working_dir 根目錄 {config.paths.working_dir} 再由程式搬入。"
                f"檔名為當日 {as_of.strftime('%Y%m%d')}.xlsx。"
                f"已檢查：{searched}。原始錯誤：{exc}"
            ) from exc
        if not config.vba.auto_closes_workbook:
            from lme_daily.excel_com import close_workbook_if_opened

            close_workbook_if_opened(workbook, opened_by_us, save_changes=False)
        else:
            logger.info(
                "不關閉參考工作簿（vba.auto_closes_workbook=true；巨集會自行關閉）"
            )

    return ready


def _execute_macro(
    app: object,
    workbook: object,
    config: AppConfig,
    *,
    prev_date: str,
    three_m_date: str,
    inputbox_prev: str,
    inputbox_three: str,
) -> None:
    names = _macro_candidates(workbook, config.vba.macro_name)
    if not config.vba.use_param_injection:
        _run_macro_with_inputboxes(
            app,
            names,
            inputbox_prev,
            inputbox_three,
            inputbox_timeout=config.vba.inputbox_timeout_seconds,
        )
        return

    last_exc: BaseException | None = None
    for name in names:
        try:
            _run_macro_with_params(app, name, prev_date, three_m_date)
            return
        except Exception as exc:
            last_exc = exc
            cause = exc.__cause__ if exc.__cause__ is not None else exc
            if is_bad_param_count(exc) or is_bad_param_count(cause):
                logger.warning(
                    "巨集 %s 不接受參數（DISP_E_BADPARAMCOUNT / 0x8002000E）。"
                    "這是正常情況（Sub 沒有參數、只用 InputBox）。"
                    "改為自動填入兩個 InputBox（格式 yyyymmdd）。",
                    name,
                )
                _run_macro_with_inputboxes(
                    app,
                    names,
                    inputbox_prev,
                    inputbox_three,
                    inputbox_timeout=config.vba.inputbox_timeout_seconds,
                )
                return
            logger.debug("參數注入 %s 失敗：%s", name, exc)

    raise ExcelComError(
        f"以參數方式執行巨集失敗（嘗試過：{names}）。Excel 錯誤：{last_exc}"
    ) from last_exc


def _macro_candidates(workbook: object, macro_name: str) -> list[str]:
    qualified = _qualified_macro_name(workbook, macro_name)
    names = [qualified]
    if macro_name not in names:
        names.append(macro_name)
    return names


def _qualified_macro_name(workbook: object, macro_name: str) -> str:
    """Excel 對含空白/中文檔名的巨集呼叫需加工作簿限定。"""
    if "!" in macro_name:
        return macro_name
    wb_name = getattr(workbook, "Name", "")
    if wb_name:
        return f"'{wb_name}'!{macro_name}"
    return macro_name


def _run_macro_with_params(app: object, macro_name: str, prev_date: str, three_m_date: str) -> None:
    logger.info("Application.Run(%s, %s, %s)", macro_name, prev_date, three_m_date)
    try:
        app.Run(macro_name, prev_date, three_m_date)
    except Exception as exc:
        wrapped = ExcelComError(f"Application.Run 參數模式失敗（{macro_name}）：{exc}")
        wrapped.__cause__ = exc
        raise wrapped from exc
    logger.info("巨集執行完畢（參數注入）")


def _run_macro_with_inputboxes(
    app: object,
    macro_names: list[str],
    prev_date: str,
    three_m_date: str,
    *,
    inputbox_timeout: float,
) -> None:
    """COM 執行緒呼叫 Run（不帶參數）；背景執行緒填兩個 InputBox。

    Excel COM 是 STA，``Application.Run`` 必須留在建立 Application 的執行緒，
    否則可能 RPC_E_WRONG_THREAD。InputBox 由 Excel 行程顯示，背景執行緒用 Win32 填入即可。
    """
    stop = threading.Event()
    fill_error: list[BaseException] = []
    filler = threading.Thread(
        target=_fill_two_inputboxes,
        kwargs={
            "prev_date": prev_date,
            "three_m_date": three_m_date,
            "timeout": inputbox_timeout,
            "stop": stop,
            "errors": fill_error,
        },
        name="lme-inputbox-filler",
        daemon=True,
    )
    filler.start()
    run_exc: BaseException | None = None
    try:
        last_exc: BaseException | None = None
        for name in macro_names:
            try:
                logger.info("Application.Run(%s)（無參數，等待 InputBox）", name)
                app.Run(name)
                logger.info("巨集執行完畢（InputBox 模式）")
                last_exc = None
                break
            except Exception as exc:
                last_exc = exc
                logger.debug("Run(%s) 失敗：%s", name, exc)
        if last_exc is not None:
            run_exc = last_exc
            raise ExcelComError(
                f"執行巨集失敗（嘗試過：{macro_names}）。請在 Excel 按 Alt+F8 確認名稱。"
                f" Excel 錯誤：{last_exc}"
            ) from last_exc
    finally:
        stop.set()
        filler.join(timeout=5)

    if fill_error and run_exc is None:
        raise ExcelComError(f"自動填 InputBox 失敗：{fill_error[0]}") from fill_error[0]


def _fill_two_inputboxes(
    *,
    prev_date: str,
    three_m_date: str,
    timeout: float,
    stop: threading.Event,
    errors: list[BaseException],
) -> None:
    try:
        _fill_one_inputbox(prev_date, timeout=timeout, which="上日日期", stop=stop)
        time.sleep(0.35)
        _fill_one_inputbox(three_m_date, timeout=timeout, which="3M date", stop=stop)
    except Exception as exc:  # noqa: BLE001
        logger.exception("填寫 InputBox 失敗")
        errors.append(exc)
        _dismiss_inputboxes()


def _fill_one_inputbox(value: str, *, timeout: float, which: str, stop: threading.Event) -> None:
    logger.info("等待 InputBox（%s），將填入 %s", which, value)
    deadline = time.monotonic() + timeout
    last_error: Exception | None = None
    while time.monotonic() < deadline and not stop.is_set():
        try:
            if _try_fill_inputbox_win32(value):
                logger.info("已用 Win32 填入 InputBox（%s）：%s", which, value)
                return
        except Exception as exc:  # noqa: BLE001
            last_error = exc
            logger.debug("Win32 InputBox：%s", exc)
        try:
            if _try_fill_inputbox_pywinauto(value):
                logger.info("已用 pywinauto 填入 InputBox（%s）：%s", which, value)
                return
        except Exception as exc:  # noqa: BLE001
            last_error = exc
            logger.debug("pywinauto InputBox：%s", exc)
        time.sleep(0.25)
    detail = f"；最後錯誤：{last_error}" if last_error else ""
    raise ExcelComError(
        f"等待 Excel InputBox（{which}）逾時（{timeout:.0f}s）。"
        "請確認巨集會彈出 InputBox，視窗未被擋住。"
        f"{detail}"
    )


def _try_fill_inputbox_win32(value: str) -> bool:
    try:
        import win32con  # type: ignore
        import win32gui  # type: ignore
    except ImportError:
        return False

    fg = win32gui.GetForegroundWindow()
    hwnds: list[int] = []
    if fg:
        try:
            if win32gui.IsWindowVisible(fg) and win32gui.GetClassName(fg) == "#32770":
                if win32gui.FindWindowEx(fg, 0, "Edit", None):
                    hwnds.append(fg)
        except Exception:
            pass

    def _enum(hwnd: int, _: object) -> bool:
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return True
            if win32gui.GetClassName(hwnd) != "#32770":
                return True
            title = win32gui.GetWindowText(hwnd) or ""
            if title and ("Excel" not in title) and ("excel" not in title.lower()) and ("輸入" not in title):
                # 仍可能是 VBA InputBox（標題常為 Microsoft Excel）；無標題也收
                if title not in {"Microsoft Excel", "Excel"}:
                    # 中文 Excel 有時標題就是提示文字，只要有 Edit 就嘗試
                    pass
            edit = win32gui.FindWindowEx(hwnd, 0, "Edit", None)
            if edit:
                hwnds.append(hwnd)
        except Exception:
            return True
        return True

    win32gui.EnumWindows(_enum, None)
    # 去重並讓前景視窗優先
    ordered: list[int] = []
    for h in hwnds:
        if h not in ordered:
            ordered.append(h)
    hwnds = ordered
    if not hwnds:
        return False

    hwnd = hwnds[0]
    edit = win32gui.FindWindowEx(hwnd, 0, "Edit", None)
    if not edit:
        return False
    win32gui.SendMessage(edit, win32con.WM_SETTEXT, 0, value)
    time.sleep(0.12)
    clicked = False
    for caption in ("確定", "OK", "Ok", "O.K."):
        btn = win32gui.FindWindowEx(hwnd, 0, "Button", caption)
        if btn:
            win32gui.SendMessage(btn, win32con.BM_CLICK, 0, 0)
            clicked = True
            break
    if not clicked:
        ok = win32gui.GetDlgItem(hwnd, 1)  # IDOK
        if ok:
            win32gui.SendMessage(ok, win32con.BM_CLICK, 0, 0)
            clicked = True
    if not clicked:
        win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
        win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
    time.sleep(0.2)
    return True


def _dismiss_inputboxes() -> None:
    """填失敗時關掉對話框，讓卡住的 Application.Run 能返回。"""
    try:
        import win32con  # type: ignore
        import win32gui  # type: ignore
    except ImportError:
        return

    def _enum(hwnd: int, _: object) -> bool:
        try:
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetClassName(hwnd) == "#32770":
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_ESCAPE, 0)
        except Exception:
            return True
        return True

    try:
        win32gui.EnumWindows(_enum, None)
    except Exception:
        return


def _try_fill_inputbox_pywinauto(value: str) -> bool:
    try:
        from pywinauto import Desktop  # type: ignore
    except ImportError:
        return False
    desktop = Desktop(backend="win32")
    dialog = _find_excel_inputbox(desktop)
    if dialog is None:
        return False
    dialog.wait("visible", timeout=3)
    edit = dialog.child_window(class_name="Edit")
    edit.wait("ready", timeout=3)
    edit.set_edit_text(value)
    try:
        dialog.type_keys("{ENTER}")
    except Exception:
        try:
            dialog.child_window(title="確定").click_input()
        except Exception:
            try:
                dialog.child_window(title="OK").click_input()
            except Exception:
                dialog.type_keys("{ENTER}")
    time.sleep(0.2)
    return True


def _find_excel_inputbox(desktop: object):
    specs = [
        {"title": "Microsoft Excel", "class_name": "#32770"},
        {"title_re": r".*Microsoft Excel.*", "class_name": "#32770"},
        {"title_re": r".*Excel.*", "class_name": "#32770"},
        {"class_name": "#32770"},
    ]
    for spec in specs:
        try:
            window = desktop.window(**spec)  # type: ignore[union-attr]
            if window.exists(timeout=0.15):
                return window
        except Exception:
            continue
    return None
