"""讀取已開啟的 LME BBG WORKBOOK（不 Open、不 Close）。

三種來源（config.bloomberg.source），都**不會**自動解鎖 Terminal：

- ``excel``：沿用使用者已手動打開的 BBG 工作簿，RefreshAll（預設）
- ``cached``：只讀目前儲存格值，不 Refresh
- ``blpapi``：Python Desktop API，不經 Excel Bloomberg 外掛

請先手動打開「LME BBG WORKBOOK.xlsx」。腳本只會讀取，不會幫你開、也不會關。
"""

from __future__ import annotations

import logging
import re
import time
from typing import Any

from lme_daily.config import AppConfig
from lme_daily.excel_com import (
    excel_app,
    wait_until_calculation_done,
)
from lme_daily.exceptions import BbgWorkbookNotOpenError, ExcelComError

logger = logging.getLogger(__name__)

RangeValues = tuple[tuple[Any, ...], ...]
RangeFormats = tuple[tuple[str, ...], ...]

BDP_RE = re.compile(
    r'(?:WSS\.)?BDP\(\s*"([^"]+)"\s*,\s*"([^"]+)"',
    re.IGNORECASE,
)


def normalize_bbg_cell(value: Any) -> Any:
    """只正規化以 N/A 開頭的字串；空白/None 維持原樣，不填 N/A。"""
    if isinstance(value, str) and value.strip().upper().startswith("N/A"):
        return "N/A"
    return value


def normalize_bbg_values(values: RangeValues) -> RangeValues:
    """讀值之後、寫入 BBG快照之前套用一次。"""
    return tuple(tuple(normalize_bbg_cell(cell) for cell in row) for row in values)


def parse_bdp_formula(cell: Any) -> tuple[str, str] | None:
    """從 ``=BDP("ticker","FIELD")`` 抽出 (security, field)。"""
    if not isinstance(cell, str):
        return None
    text = cell.strip()
    if text.startswith("="):
        text = text[1:]
    match = BDP_RE.search(text)
    if not match:
        return None
    return match.group(1).strip(), match.group(2).strip()


def fetch_bloomberg_snapshot(config: AppConfig) -> tuple[RangeValues, RangeFormats]:
    source = config.bloomberg.source
    logger.info("Bloomberg 取數來源：%s", source)
    if source == "blpapi":
        from lme_daily.bbg_blpapi import fetch_bloomberg_snapshot_blpapi

        values, formats = fetch_bloomberg_snapshot_blpapi(config)
    else:
        refresh = source != "cached"
        values, formats = _fetch_via_excel(config, refresh=refresh)
    return normalize_bbg_values(values), formats


def _attach_open_bbg_workbook(app: object, config: AppConfig) -> Any:
    """只抓使用者已手動打開的 BBG 工作簿；絕不 Workbooks.Open。"""
    name = config.paths.bbg_workbook_name
    try:
        workbook = app.Workbooks(name)  # type: ignore[attr-defined]
    except Exception as exc:
        raise BbgWorkbookNotOpenError(
            f"找不到已開啟的「{name}」，請先手動打開這個工作簿再重新執行。"
        ) from exc
    logger.info("沿用已開啟的 BBG 工作簿：%s（不 Open、不 Close、不 Activate）", name)
    return workbook


def _fetch_via_excel(config: AppConfig, *, refresh: bool) -> tuple[RangeValues, RangeFormats]:
    sheet_name = config.bloomberg.bbg_sheet_name
    copy_range = config.bloomberg.copy_range
    logger.info(
        "Excel 讀取 BBG：sheet=%s range=%s refresh=%s wait=%ss",
        sheet_name,
        copy_range,
        refresh,
        config.bloomberg.refresh_wait_seconds,
    )

    with excel_app(
        visible=config.excel.visible,
        display_alerts=config.excel.display_alerts,
        reuse_running=config.excel.reuse_running,
        quit_on_exit=config.excel.quit_on_exit,
        new_instance=config.excel.new_instance,
    ) as app:
        workbook = _attach_open_bbg_workbook(app, config)
        if refresh:
            _refresh_bloomberg(app, workbook, config)
        else:
            logger.info("cached 模式：不呼叫 RefreshAll")
        values, formats = _read_range(workbook, sheet_name, copy_range)

    logger.info("已讀取 BBG 區間 %s!%s（%d 列）", sheet_name, copy_range, len(values))
    return values, formats


def _refresh_bloomberg(app: object, workbook: object, config: AppConfig) -> None:
    try:
        workbook.RefreshAll()
        logger.info("已呼叫 Workbook.RefreshAll()")
    except Exception as exc:
        raise ExcelComError(f"RefreshAll 失敗：{exc}") from exc

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
