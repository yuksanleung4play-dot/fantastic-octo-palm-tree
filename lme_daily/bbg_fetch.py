"""開啟 LME BBG WORKBOOK，刷新 Bloomberg 連結並讀取指定區間的值。"""

from __future__ import annotations

import logging
import time
from typing import Any

from lme_daily.config import AppConfig
from lme_daily.excel_com import (
    close_workbook,
    excel_app,
    open_workbook,
    wait_until_calculation_done,
)
from lme_daily.exceptions import ExcelComError

logger = logging.getLogger(__name__)

RangeValues = tuple[tuple[Any, ...], ...]
RangeFormats = tuple[tuple[str, ...], ...]


def fetch_bloomberg_snapshot(config: AppConfig) -> tuple[RangeValues, RangeFormats]:
    """RefreshAll → 等待 → 以 Value2 讀取 copy_range，然後不儲存關閉。"""
    sheet_name = config.bloomberg.bbg_sheet_name
    copy_range = config.bloomberg.copy_range
    logger.info(
        "開啟 BBG 工作簿並刷新：sheet=%s range=%s wait=%ss",
        sheet_name,
        copy_range,
        config.bloomberg.refresh_wait_seconds,
    )

    with excel_app(visible=config.excel.visible, display_alerts=config.excel.display_alerts) as app:
        workbook = open_workbook(app, config.paths.bbg_workbook, update_links=3)
        try:
            _refresh_bloomberg(app, workbook, config)
            values, formats = _read_range(workbook, sheet_name, copy_range)
        finally:
            close_workbook(workbook, save_changes=False)

    logger.info("已讀取 BBG 區間 %s!%s（%d 列）", sheet_name, copy_range, len(values))
    return values, formats


def _refresh_bloomberg(app: object, workbook: object, config: AppConfig) -> None:
    try:
        workbook.RefreshAll()
        logger.info("已呼叫 Workbook.RefreshAll()")
    except Exception as exc:
        raise ExcelComError(f"RefreshAll 失敗：{exc}") from exc

    # Bloomberg 常用 RTD / 非同步查詢；Excel 2013+ 提供此方法
    try:
        if hasattr(app, "CalculateUntilAsyncQueriesDone"):
            logger.info("呼叫 CalculateUntilAsyncQueriesDone()")
            app.CalculateUntilAsyncQueriesDone()
    except Exception as exc:
        logger.warning("CalculateUntilAsyncQueriesDone 失敗（將改以 sleep + CalculationState）：%s", exc)

    wait = config.bloomberg.refresh_wait_seconds
    logger.info("等待 Bloomberg 刷新 %.1f 秒", wait)
    time.sleep(wait)
    wait_until_calculation_done(
        app,
        timeout_seconds=config.bloomberg.calculation_timeout_seconds,
    )


def _read_range(
    workbook: object,
    sheet_name: str,
    copy_range: str,
) -> tuple[RangeValues, RangeFormats]:
    try:
        sheet = workbook.Worksheets(sheet_name)
    except Exception as exc:
        raise ExcelComError(
            f"找不到工作表 {sheet_name!r}。請核對 config.bloomberg.bbg_sheet_name。錯誤：{exc}"
        ) from exc

    try:
        rng = sheet.Range(copy_range)
        raw_values = rng.Value2
        raw_formats = rng.NumberFormat
    except Exception as exc:
        raise ExcelComError(f"讀取 {sheet_name}!{copy_range} 失敗：{exc}") from exc

    values = _as_2d_tuple(raw_values)
    formats = _as_2d_tuple(raw_formats, as_str=True)
    if not values:
        raise ExcelComError(f"{sheet_name}!{copy_range} 讀到空值（Value2 為 None）")
    return values, formats


def _as_2d_tuple(raw: Any, *, as_str: bool = False) -> tuple[tuple[Any, ...], ...]:
    """把 COM Range.Value2 正規化成 2D tuple（單格也包成 1x1）。"""
    if raw is None:
        return ()
    if not isinstance(raw, tuple):
        cell = str(raw) if as_str else raw
        return ((cell,),)
    if not raw:
        return ()
    first = raw[0]
    if not isinstance(first, tuple):
        row = tuple(str(v) if as_str else v for v in raw)
        return (row,)
    rows = []
    for row in raw:
        rows.append(tuple(str(v) if as_str else v for v in row))
    return tuple(rows)
