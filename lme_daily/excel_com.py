"""Excel COM 共用封裝（Windows + 本機 Excel）。

預設沿用已開啟的 Excel，結束時不 Quit。DispatchEx + Quit 會拆掉 Bloomberg
Excel 外掛連線，Terminal 常被自動鎖上。
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
    """取得 Excel.Application。

    - ``reuse_running=True``：優先 GetActiveObject，沿用你已登入 Bloomberg 的 Excel。
    - ``new_instance=False``：不要 DispatchEx（獨立行程容易讓 Terminal 上鎖）。
    - ``quit_on_exit=False``：離開時不 Quit，只關我們開啟的工作簿。
    """
    win32com_client, pythoncom = import_win32com()
    pythoncom.CoInitialize()
    app = None
    started_by_us = False
    try:
        app, started_by_us = _acquire_excel(
            win32com_client,
            reuse_running=reuse_running,
            new_instance=new_instance,
        )
        app.Visible = visible
        app.DisplayAlerts = display_alerts
        try:
            app.AskToUpdateLinks = False
        except Exception:
            logger.debug("AskToUpdateLinks 無法設定，略過")
        yield app
    except ExcelComError:
        raise
    except Exception as exc:
        raise ExcelComError(f"啟動 Excel 失敗：{exc}") from exc
    finally:
        if app is not None and quit_on_exit and started_by_us:
            try:
                app.DisplayAlerts = False
                app.Quit()
                logger.info("已關閉 Excel.Application（quit_on_exit=true 且由本程式啟動）")
            except Exception as exc:
                logger.warning("Excel.Quit 失敗：%s", exc)
        elif app is not None:
            logger.info("Excel 保持開啟，不呼叫 Quit（避免 Bloomberg Terminal 被鎖）")
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _acquire_excel(win32com_client: Any, *, reuse_running: bool, new_instance: bool) -> tuple[Any, bool]:
    if reuse_running and not new_instance:
        try:
            app = win32com_client.GetActiveObject("Excel.Application")
            logger.info("沿用已開啟的 Excel.Application（GetActiveObject）")
            return app, False
        except Exception:
            logger.info("沒有已開啟的 Excel，改為 Dispatch 啟動（會載入 Bloomberg 外掛）")

    if new_instance:
        logger.warning(
            "excel.new_instance=true 使用 DispatchEx。獨立 Excel 行程常導致 Bloomberg Terminal 上鎖。"
        )
        return win32com_client.DispatchEx("Excel.Application"), True

    return win32com_client.Dispatch("Excel.Application"), True


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
    if not resolved.is_file():
        raise ExcelComError(f"工作簿不存在：{resolved}")
    logger.info("開啟工作簿：%s", resolved)
    try:
        workbook = app.Workbooks.Open(
            str(resolved),
            UpdateLinks=update_links,
            ReadOnly=read_only,
        )
        return workbook, True
    except Exception as exc:
        raise ExcelComError(f"開啟工作簿失敗（{resolved}）：{exc}") from exc


def close_workbook(workbook: Any, *, save_changes: bool = False) -> None:
    name = getattr(workbook, "Name", "<unknown>")
    try:
        workbook.Close(SaveChanges=save_changes)
        logger.info("已關閉工作簿：%s（SaveChanges=%s）", name, save_changes)
    except Exception as exc:
        raise ExcelComError(f"關閉工作簿失敗（{name}）：{exc}") from exc


def close_workbook_if_opened(workbook: Any, opened_by_us: bool, *, save_changes: bool = False) -> None:
    if opened_by_us:
        close_workbook(workbook, save_changes=save_changes)
    else:
        logger.info(
            "工作簿原本就開著，不關閉：%s",
            getattr(workbook, "Name", "<unknown>"),
        )


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
    deadline = time.monotonic() + timeout_seconds
    logger.info("等待輸出檔產生：%s（timeout=%.0fs）", path, timeout_seconds)
    while time.monotonic() < deadline:
        if path.exists() and path.stat().st_size > 0:
            if _file_unlocked(path):
                logger.info("輸出檔已就緒：%s（%d bytes）", path, path.stat().st_size)
                return path
            logger.debug("檔案存在但仍被鎖定，繼續等待：%s", path)
        time.sleep(poll_interval)
    raise ExcelComError(
        f"等待檔案逾時（{timeout_seconds:.0f}s）仍未產生或無法讀取：{path}"
    )


def _file_unlocked(path: Path) -> bool:
    try:
        with path.open("rb"):
            return True
    except OSError:
        return False
