"""Excel COM 共用封裝（Windows + 本機 Excel）。

Linux / 無 pywin32 環境仍可 import 本模組；真正呼叫 COM 時才會失敗並給出明確訊息。
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

# Excel XlCalculationState
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
def excel_app(*, visible: bool = True, display_alerts: bool = False) -> Iterator[Any]:
    """啟動獨立 Excel 執行個體（DispatchEx），結束時關閉。"""
    win32com_client, pythoncom = import_win32com()
    pythoncom.CoInitialize()
    app = None
    try:
        logger.info("啟動 Excel.Application（DispatchEx, Visible=%s）", visible)
        app = win32com_client.DispatchEx("Excel.Application")
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
        if app is not None:
            try:
                app.DisplayAlerts = False
                app.Quit()
                logger.info("已關閉 Excel.Application")
            except Exception as exc:
                logger.warning("Excel.Quit 失敗：%s", exc)
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def open_workbook(app: Any, path: Path, *, read_only: bool = False, update_links: int = 0) -> Any:
    resolved = path.resolve()
    if not resolved.is_file():
        raise ExcelComError(f"工作簿不存在：{resolved}")
    logger.info("開啟工作簿：%s", resolved)
    try:
        return app.Workbooks.Open(
            str(resolved),
            UpdateLinks=update_links,
            ReadOnly=read_only,
        )
    except Exception as exc:
        raise ExcelComError(f"開啟工作簿失敗（{resolved}）：{exc}") from exc


def close_workbook(workbook: Any, *, save_changes: bool = False) -> None:
    name = getattr(workbook, "Name", "<unknown>")
    try:
        workbook.Close(SaveChanges=save_changes)
        logger.info("已關閉工作簿：%s（SaveChanges=%s）", name, save_changes)
    except Exception as exc:
        raise ExcelComError(f"關閉工作簿失敗（{name}）：{exc}") from exc


def wait_until_calculation_done(
    app: Any,
    *,
    timeout_seconds: float,
    poll_interval: float = 0.5,
) -> None:
    """等到 Application.CalculationState == 0（xlDone）。"""
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
    """Polling 直到檔案存在且大小 > 0；Windows 上再嘗試確認未被鎖定。"""
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
