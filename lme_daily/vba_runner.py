"""開啟早班 LME reference 工作簿並執行 VBA 巨集。

兩種填入「上日日期 / 3M date」的方式，由 config ``vba.use_param_injection`` 切換：

1. **參數注入（優先）**：``Application.Run(macro_name, prev_date, three_m_date)``
   前提是巨集宣告 Optional 參數並在缺失時才呼叫 InputBox。範例見
   ``examples/RunDailyLME_param_wrapper.bas``。
2. **pywinauto 備援**：巨集仍用 InputBox 時，在背景執行巨集並自動填兩個彈窗。

TODO: 請確認 ``vba.macro_name`` 的真實 Sub 名稱，以及巨集是否支援參數覆蓋 InputBox。
若不支援，請把 ``use_param_injection`` 設為 false。
"""

from __future__ import annotations

import logging
import threading
import time
from datetime import date
from pathlib import Path

from lme_daily.config import AppConfig
from lme_daily.excel_com import (
    close_workbook,
    excel_app,
    import_win32com,
    open_workbook,
    wait_for_file,
)
from lme_daily.exceptions import ExcelComError, MacroOutputError

logger = logging.getLogger(__name__)


def run_reference_macro(
    config: AppConfig,
    *,
    as_of: date,
    prev_date: str,
    three_m_date: str,
) -> Path:
    """執行巨集、等到 ``yyyymmdd.xlsx`` 產生後關閉參考工作簿。"""
    expected = config.step2_workbook_path(as_of)
    if expected.exists():
        logger.warning("輸出檔已存在，將覆寫：%s", expected)

    logger.info(
        "VBA 模式：%s；巨集=%s；上日=%s；3M=%s",
        "參數注入" if config.vba.use_param_injection else "pywinauto InputBox",
        config.vba.macro_name,
        prev_date,
        three_m_date,
    )

    ready = expected
    with excel_app(visible=config.excel.visible, display_alerts=config.excel.display_alerts) as app:
        workbook = open_workbook(app, config.paths.ref_workbook)
        try:
            if config.vba.use_param_injection:
                _run_macro_with_params(app, workbook, config.vba.macro_name, prev_date, three_m_date)
            else:
                _run_macro_with_pywinauto(
                    app,
                    workbook,
                    config.vba.macro_name,
                    prev_date,
                    three_m_date,
                    inputbox_timeout=config.vba.inputbox_timeout_seconds,
                )
            try:
                ready = wait_for_file(
                    expected,
                    timeout_seconds=config.vba.output_timeout_seconds,
                    poll_interval=config.vba.poll_interval_seconds,
                )
            except ExcelComError as exc:
                raise MacroOutputError(
                    f"巨集執行後未產生預期檔案 {expected.name}。"
                    f"請確認巨集會把結果寫到 {config.paths.working_dir}，"
                    f"且檔名為當日 {as_of.strftime('%Y%m%d')}.xlsx。原始錯誤：{exc}"
                ) from exc
        except ExcelComError:
            raise
        except MacroOutputError:
            raise
        except Exception as exc:
            raise ExcelComError(f"執行巨集 {config.vba.macro_name!r} 失敗：{exc}") from exc
        finally:
            close_workbook(workbook, save_changes=False)

    return ready


def _qualified_macro_name(workbook: object, macro_name: str) -> str:
    """Excel 對含空白/中文檔名的巨集呼叫需加工作簿限定。"""
    if "!" in macro_name:
        return macro_name
    wb_name = getattr(workbook, "Name", "")
    if wb_name:
        return f"'{wb_name}'!{macro_name}"
    return macro_name


def _run_macro_with_params(
    app: object,
    workbook: object,
    macro_name: str,
    prev_date: str,
    three_m_date: str,
) -> None:
    qualified = _qualified_macro_name(workbook, macro_name)
    logger.info("Application.Run(%s, %s, %s)", qualified, prev_date, three_m_date)
    # TODO: 若巨集不接受參數，這裡會失敗——改 config.vba.use_param_injection=false
    try:
        app.Run(qualified, prev_date, three_m_date)
    except Exception as exc:
        raise ExcelComError(
            f"以參數方式執行巨集失敗（{qualified}）。"
            "若巨集只有 InputBox、沒有 Optional 參數，請將 "
            "vba.use_param_injection 設為 false 改走 pywinauto。"
            f" Excel 錯誤：{exc}"
        ) from exc
    logger.info("巨集執行完畢（參數注入）")


def _run_macro_with_pywinauto(
    app: object,
    workbook: object,
    macro_name: str,
    prev_date: str,
    three_m_date: str,
    *,
    inputbox_timeout: float,
) -> None:
    try:
        from pywinauto import Desktop  # type: ignore
    except ImportError as exc:
        raise ExcelComError(
            "use_param_injection=false 需要 pywinauto。請執行：pip install pywinauto"
        ) from exc

    _, pythoncom = import_win32com()
    qualified = _qualified_macro_name(workbook, macro_name)
    errors: list[BaseException] = []
    finished = threading.Event()

    def _invoke() -> None:
        pythoncom.CoInitialize()
        try:
            logger.info("背景執行 Application.Run(%s) 並等待 InputBox", qualified)
            app.Run(qualified)
            logger.info("巨集執行完畢（pywinauto 模式）")
        except Exception as exc:  # noqa: BLE001 — 傳到主執行緒再轉成 ExcelComError
            errors.append(exc)
        finally:
            finished.set()
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    worker = threading.Thread(target=_invoke, name="lme-vba-run", daemon=True)
    worker.start()
    try:
        _fill_inputbox(Desktop, prev_date, timeout=inputbox_timeout, which="上日日期")
        time.sleep(0.4)
        _fill_inputbox(Desktop, three_m_date, timeout=inputbox_timeout, which="3M date")
    except Exception:
        logger.exception("填寫 InputBox 失敗")
        raise

    if not finished.wait(timeout=max(inputbox_timeout * 2, 30)):
        raise ExcelComError("巨集在填完 InputBox 後仍未結束，請檢查 VBA 是否卡住。")
    worker.join(timeout=5)
    if errors:
        raise ExcelComError(f"巨集執行例外：{errors[0]}") from errors[0]


def _fill_inputbox(desktop_cls: object, value: str, *, timeout: float, which: str) -> None:
    """尋找 Excel InputBox（標題通常為 Microsoft Excel）並填入後按 Enter。"""
    logger.info("等待 InputBox（%s），將填入 %s", which, value)
    desktop = desktop_cls(backend="win32")  # type: ignore[operator]
    deadline = time.monotonic() + timeout
    last_error: Exception | None = None
    while time.monotonic() < deadline:
        dialog = _find_excel_inputbox(desktop)
        if dialog is not None:
            try:
                dialog.wait("visible", timeout=5)
                edit = dialog.child_window(class_name="Edit")
                edit.wait("ready", timeout=5)
                edit.set_edit_text(value)
                logger.info("已在 InputBox（%s）填入：%s", which, value)
                try:
                    dialog.type_keys("{ENTER}")
                except Exception:
                    try:
                        dialog.child_window(title="OK").click_input()
                    except Exception:
                        dialog.type_keys("{ENTER}")
                time.sleep(0.3)
                return
            except Exception as exc:  # noqa: BLE001
                last_error = exc
                logger.debug("InputBox 尚未就緒：%s", exc)
        time.sleep(0.4)
    detail = f"；最後錯誤：{last_error}" if last_error else ""
    raise ExcelComError(
        f"等待 Excel InputBox（{which}）逾時（{timeout:.0f}s）。"
        "請確認巨集確實會彈出 InputBox，且 Excel 視窗未被其他程式擋住。"
        f"{detail}"
    )


def _find_excel_inputbox(desktop: object):
    """嘗試以常見標題 / class 找到 InputBox 對話框。"""
    specs = [
        {"title": "Microsoft Excel", "class_name": "#32770"},
        {"title_re": r".*Microsoft Excel.*", "class_name": "#32770"},
        {"title_re": r".*Excel.*", "class_name": "#32770"},
    ]
    for spec in specs:
        try:
            window = desktop.window(**spec)  # type: ignore[union-attr]
            if window.exists(timeout=0.2):
                return window
        except Exception:
            continue
    return None
