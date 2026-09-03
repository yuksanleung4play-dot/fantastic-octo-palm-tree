"""讀取 LME BBG WORKBOOK（已開就沿用，未開就 Open；讀完不 Close）。

三種來源（config.bloomberg.source），都**不會**自動解鎖 Terminal：

- ``excel``：沿用或開啟 BBG 工作簿，等待後 RefreshAll（預設）
- ``cached``：只讀目前儲存格值，不 Refresh
- ``blpapi``：Python Desktop API，不經 Excel Bloomberg 外掛

腳本會自動開啟這個工作簿（若尚未開啟），開啟後等待設定秒數再讀取，讀取後不會關閉。
"""

from __future__ import annotations

import logging
import re
import time
from datetime import date, datetime, timedelta
from typing import Any

from lme_daily.config import AppConfig
from lme_daily.excel_com import (
    _record_workbook_open,
    excel_app,
    wait_until_calculation_done,
)
from lme_daily.exceptions import ExcelComError

logger = logging.getLogger(__name__)

RangeValues = tuple[tuple[Any, ...], ...]
RangeFormats = tuple[tuple[str, ...], ...]

# Excel Value2 日期序列的 epoch（與 report_builder 相同）
EXCEL_EPOCH = datetime(1899, 12, 30)
_PROMPT_DATE_TEXT_FORMATS = (
    "%Y%m%d",
    "%Y/%m/%d",
    "%Y-%m-%d",
    "%d/%m/%Y",
    "%m/%d/%Y",
    "%d-%b-%y",
    "%d-%b-%Y",
    "%d/%b/%y",
    "%d/%b/%Y",
    "%b %d, %Y",
)

BDP_RE = re.compile(
    r'(?:WSS\.)?BDP\(\s*"([^"]+)"\s*,\s*"([^"]+)"',
    re.IGNORECASE,
)


def normalize_bbg_cell(value: Any) -> Any:
    """只正規化以 N/A 開頭的字串；空白/None 維持原樣，不填 N/A。"""
    if isinstance(value, str) and value.strip().upper().startswith("N/A"):
        return "N/A"
    return value


def parse_excel_display_text(text: Any) -> Any:
    """把 Excel ``cell.Text``（套用格式後的顯示字串）轉回報告用的值。

    例如 ``"16,554.09"`` → ``16554.09``，與畫面上的四捨五入結果一致。
    """
    if text is None:
        return None
    if isinstance(text, bool):
        return text
    if isinstance(text, (datetime, date)):
        return text
    if not isinstance(text, str):
        text = str(text)
    cleaned = text.replace(",", "").replace("\xa0", "").strip()
    if cleaned.upper().startswith("N/A") or cleaned.startswith("#N/A"):
        return "N/A"
    if cleaned == "":
        return None
    try:
        return float(cleaned)
    except ValueError:
        return text


def value_from_stored_number(value: Any) -> Any:
    """openpyxl / pandas 沒有 ``.Text`` 時，用兩位小數顯示字串再解析，模擬 Excel 畫面。"""
    if value is None:
        return None
    if isinstance(value, bool) or isinstance(value, (datetime, date)):
        return value
    if isinstance(value, str):
        return parse_excel_display_text(value)
    try:
        num = float(value)
    except (TypeError, ValueError):
        return parse_excel_display_text(str(value))
    if num != num:  # NaN
        return value
    return parse_excel_display_text(f"{num:,.2f}")


def normalize_bbg_values(values: RangeValues) -> RangeValues:
    """讀值之後、寫入 BBG快照之前套用一次。"""
    return tuple(tuple(normalize_bbg_cell(cell) for cell in row) for row in values)


def parse_bbg_prompt_date_value(value: Any) -> str | None:
    """把 LME_PROMPT_DT 儲存格值轉成 ``yyyymmdd``；N/A、空白、無法解析則 ``None``。"""
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, datetime):
        return value.date().strftime("%Y%m%d")
    if isinstance(value, date):
        return value.strftime("%Y%m%d")
    if isinstance(value, (int, float)):
        return _excel_serial_to_yyyymmdd(float(value))
    if not isinstance(value, str):
        return None
    cleaned = value.replace("\xa0", " ").strip()
    if cleaned == "":
        return None
    if cleaned.upper().startswith("N/A") or cleaned.startswith("#N/A"):
        return None
    try:
        numeric = float(cleaned.replace(",", ""))
    except ValueError:
        numeric = None
    if numeric is not None:
        parsed = _excel_serial_to_yyyymmdd(numeric)
        if parsed is not None:
            return parsed
    for fmt in _PROMPT_DATE_TEXT_FORMATS:
        try:
            return datetime.strptime(cleaned, fmt).strftime("%Y%m%d")
        except ValueError:
            continue
    return None


def _excel_serial_to_yyyymmdd(serial: float) -> str | None:
    if serial != serial:  # NaN
        return None
    if not 20000 < serial < 80000:
        return None
    try:
        return (EXCEL_EPOCH + timedelta(days=serial)).date().strftime("%Y%m%d")
    except (OverflowError, ValueError, OSError):
        return None


def fetch_bbg_3m_prompt_date(
    app: object,
    config: AppConfig,
    *,
    fallback: str | None = None,
    workbook: Any | None = None,
) -> str | None:
    """從 BBG 工作簿讀 ``bloomberg.prompt_date_cell``（LME_PROMPT_DT）成 yyyymmdd。

    優先 ``Value2``（Excel 日期序列），再試 ``Text``。N/A / 空白 / 無法解析時
    回傳 ``fallback``（Python ``calc_lme_dates`` 的 3M）並記 WARNING。
    """
    cell_addr = (config.bloomberg.prompt_date_cell or "B4").strip() or "B4"
    sheet_name = config.bloomberg.bbg_sheet_name
    raw_value2: Any = None
    raw_text: Any = None
    try:
        workbook = workbook or _attach_open_bbg_workbook(app, config)
        sheet = workbook.Worksheets(sheet_name)
        cell = sheet.Range(cell_addr)
        try:
            raw_value2 = cell.Value2
        except Exception:
            raw_value2 = None
        parsed = parse_bbg_prompt_date_value(raw_value2)
        if parsed is None:
            try:
                raw_text = cell.Text
            except Exception:
                raw_text = None
            parsed = parse_bbg_prompt_date_value(raw_text)
    except Exception as exc:
        logger.warning(
            "讀取 BBG 3M Prompt Date（%s!%s）失敗：%s；改用 Python 計算的 3M date",
            sheet_name,
            cell_addr,
            exc,
        )
        return fallback

    if parsed is not None:
        logger.info(
            "BBG 3M Prompt Date（%s!%s LME_PROMPT_DT）= %s（Value2=%r Text=%r）",
            sheet_name,
            cell_addr,
            parsed,
            raw_value2,
            raw_text,
        )
        return parsed

    logger.warning(
        "BBG 3M Prompt Date（%s!%s）無法解析（Value2=%r Text=%r）；改用 Python 計算的 3M date %s",
        sheet_name,
        cell_addr,
        raw_value2,
        raw_text,
        fallback,
    )
    return fallback


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


def fetch_bloomberg_snapshot(
    config: AppConfig,
    *,
    fallback_3m: str | None = None,
) -> tuple[RangeValues, RangeFormats]:
    values, formats, _bbg_3m = fetch_bloomberg_snapshot_and_3m(config, fallback_3m=fallback_3m)
    return values, formats


def fetch_bloomberg_snapshot_and_3m(
    config: AppConfig,
    *,
    fallback_3m: str | None = None,
) -> tuple[RangeValues, RangeFormats, str | None]:
    """刷新並讀 copy_range，同時讀 LME_PROMPT_DT 成 yyyymmdd（失敗則 fallback）。"""
    source = config.bloomberg.source
    logger.info("Bloomberg 取數來源：%s", source)
    bbg_3m: str | None = fallback_3m
    if source == "blpapi":
        from lme_daily.bbg_blpapi import fetch_bloomberg_snapshot_blpapi

        values, formats = fetch_bloomberg_snapshot_blpapi(config)
        if fallback_3m:
            logger.warning(
                "blpapi 模式沒有 Excel LME_PROMPT_DT 儲存格，3M date 使用 Python 計算值 %s",
                fallback_3m,
            )
        bbg_3m = fallback_3m
    else:
        refresh = source != "cached"
        values, formats, bbg_3m = _fetch_via_excel(
            config, refresh=refresh, fallback_3m=fallback_3m
        )
    return normalize_bbg_values(values), formats, bbg_3m


def format_3m_date_for_vba(ymd: str, date_format: str) -> str:
    """把 yyyymmdd 轉成巨集參數 / InputBox 所需格式。"""
    return datetime.strptime(ymd, "%Y%m%d").strftime(date_format)


def _attach_open_bbg_workbook(app: object, config: AppConfig) -> Any:
    """已開著就沿用；抓不到再 ``Workbooks.Open``。讀完由呼叫端決定不 Close。"""
    name = config.paths.bbg_workbook_name
    try:
        workbook = app.Workbooks(name)  # type: ignore[attr-defined]
        logger.info("沿用已開啟的 %s", name)
        return workbook
    except Exception as exc:
        logger.info("%s 未開啟，改由程式開啟", name)
        logger.debug("Workbooks(%s) 失敗：%s", name, exc)
    return _open_bbg_workbook(app, config)


def _open_bbg_workbook(app: object, config: AppConfig) -> Any:
    path = config.paths.bbg_workbook
    try:
        app.DisplayAlerts = False  # type: ignore[attr-defined]
    except Exception:
        logger.debug("DisplayAlerts 無法設定，繼續 Open")
    _record_workbook_open()
    try:
        workbook = app.Workbooks.Open(  # type: ignore[attr-defined]
            str(path),
            UpdateLinks=0,
            IgnoreReadOnlyRecommended=True,
        )
    except Exception as exc:
        raise ExcelComError(f"開啟 Bloomberg 工作簿失敗（{path}）：{exc}") from exc
    logger.info("已用程式開啟：%s（讀取後不 Close）", path)
    return workbook


def _fetch_via_excel(
    config: AppConfig,
    *,
    refresh: bool,
    fallback_3m: str | None = None,
) -> tuple[RangeValues, RangeFormats, str | None]:
    sheet_name = config.bloomberg.bbg_sheet_name
    copy_range = config.bloomberg.copy_range
    wait = config.bloomberg.refresh_wait_seconds
    logger.info(
        "Excel 讀取 BBG：sheet=%s range=%s prompt_date_cell=%s refresh=%s wait=%ss",
        sheet_name,
        copy_range,
        config.bloomberg.prompt_date_cell,
        refresh,
        wait,
    )

    with excel_app(
        visible=config.excel.visible,
        display_alerts=config.excel.display_alerts,
        reuse_running=config.excel.reuse_running,
        quit_on_exit=config.excel.quit_on_exit,
        new_instance=config.excel.new_instance,
    ) as app:
        workbook = _attach_open_bbg_workbook(app, config)
        logger.info("已開啟，等待 %s 秒後開始讀取", wait)
        time.sleep(wait)
        if refresh:
            _refresh_bloomberg(app, workbook, config)
        else:
            logger.info("cached 模式：不呼叫 RefreshAll")
        values, formats = _read_range(workbook, sheet_name, copy_range)
        bbg_3m = fetch_bbg_3m_prompt_date(
            app, config, fallback=fallback_3m, workbook=workbook
        )

    logger.info("已讀取 BBG 區間 %s!%s（%d 列）", sheet_name, copy_range, len(values))
    return values, formats, bbg_3m


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
        logger.warning("CalculateUntilAsyncQueriesDone 失敗（將改以 CalculationState）：%s", exc)

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
        raw_formats = rng.NumberFormat
        n_rows = int(rng.Rows.Count)
        n_cols = int(rng.Columns.Count)
        parsed_rows: list[tuple[Any, ...]] = []
        for r in range(1, n_rows + 1):
            row_vals: list[Any] = []
            for c in range(1, n_cols + 1):
                cell = rng.Cells(r, c)
                try:
                    text = cell.Text
                except Exception:
                    text = getattr(cell, "Value2", None)
                row_vals.append(parse_excel_display_text(text))
            parsed_rows.append(tuple(row_vals))
        values = tuple(parsed_rows)
    except Exception as exc:
        raise ExcelComError(f"讀取 {sheet_name}!{copy_range} 失敗：{exc}") from exc

    formats = _as_2d_tuple(raw_formats, as_str=True)
    if not values:
        raise ExcelComError(f"{sheet_name}!{copy_range} 讀到空值")
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
