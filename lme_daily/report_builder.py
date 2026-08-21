"""產生最終報告 Excel：BBG快照 / 遠期走勢圖 / 原始數據 / 合併版面。

圖表引擎由 ``chart.engine`` 切換：
- ``matplotlib``：靜態 PNG 嵌入 openpyxl（預設）
- ``xlsxwriter``：Excel 原生可互動折線圖

「遠期走勢圖」與「原始數據」的資料來源一律是 ``vba_dir / yyyymmdd.xlsx``，
不跟隨 ``output_dir`` / ``run_dir``。原始數據 sheet 貼入靜態值，不使用外部連結。
"""

from __future__ import annotations

import logging
import shutil
import tempfile
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import pandas as pd
from dateutil.relativedelta import relativedelta
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.worksheet import Worksheet

from lme_daily.bbg_fetch import normalize_bbg_cell, value_from_stored_number
from lme_daily.config import AppConfig
from lme_daily.exceptions import ReportBuildError

logger = logging.getLogger(__name__)

DATE_COL = "Prompt"
SIX_SERIES = {
    "CA": "銅 Copper",
    "AH": "鋁 Aluminium",
    "ZS": "鋅 Zinc",
    "NI": "鎳 Nickel",
    "SN": "錫 Tin",
    "PB": "鉛 Lead",
}

SHEET_BBG = "BBG快照"
SHEET_CHART = "遠期走勢圖"
SHEET_RAW = "原始數據"
SHEET_PRINT = "合併版面"
SHEET_CHART_DATA = "_chart_data"
PRINT_RAW_SECTION_TITLE = "Prompt date settlement price"

# BBG快照 / 遠期走勢圖 sheet 仍用深藍，避免動到非合併版面頁。
COLOR_NAVY = "0B2447"
COLOR_GOLD = "C5A059"
COLOR_GRAY = "F2F4F7"
COLOR_LINE = "D6D9DE"
COLOR_MUTED = "6B7280"
COLOR_INK = "1F2937"
COLOR_WHITE = "FFFFFF"
# 合併版面：山證國際官網暖色（磚紅／橙漸層替代），白底。
COLOR_BRICK = "A63A2E"
COLOR_CREAM = "FFF5E8"
COLOR_ACCENT = "C9683A"  # #B5502C→#F2CDA8 橫向漸層的實色替代
COLOR_ALERT = "D9291C"
COLOR_TITLE_DARK = "2B2B2B"

PRINT_FOOTER = "第 &P 頁，共 &N 頁"
PRINT_DISCLAIMER_FULL = (
    "免責聲明：本報告由山證國際金融控股有限公司（\"本公司\"或\"山證國際\"）提供，"
    "所載之內容或意見乃根據本公司認為可靠之數據源來編制，惟本公司並不就此等內容之準確性、"
    "完整性及正確性作出明示或默示之保證。本報告的作用純粹為提供信息，並不應視為對本報告內"
    "提及的任何產品買賣或交易之專業推介、建議、邀請或要約。"
    "\n\n"
    "投資附帶風險，投資者需注意投資項目之價值可升亦可跌，而過往之表現亦不一定反映未來之表現。"
    "投資者進行投資前請尋求獨立之投資意見。本公司竭力確保其提供之數據準確可靠，但不保證該等"
    "數據絕對正確可靠；對於任何因資料不確或遺漏又或因根據或倚賴本資料所作決定、行動或不行動"
    "而引致之損失或損害，本公司概不負責（不論是民事侵權行為責任或合約責任或其他）。"
)
DISCLAIMER_FONT_SIZE = 8
DISCLAIMER_ROW_HEIGHT = 96
CHART_ROW_STRIDE = 19
PAPERSIZE_A3 = 8
PAPERSIZE_A4 = 9
# Excel 頁首/頁尾 ``&nn`` 必須兩位數字。``&8`` 接 ``2026-08-20`` 會變成 ``&82``（82pt）。
HF_FONT_SIZE = "08"

EXCEL_EPOCH = datetime(1899, 12, 30)
SOURCE_LABEL = "資料來源"

COMPANY_ZH = "山證國際金融控股有限公司"
COMPANY_EN = "Shanxi Securities International"
REPORT_TITLE_ZH = "LME每日報價"
REPORT_TITLE_EN = "LME Daily Quotation"
NUMBER_FORMAT_2DP = "#,##0.00"
AUTOFIT_PADDING = 3
# 下限：容納 5 位數 + 小數點 + 2 位小數（如 16554.08 / 16,554.08）；不是上限。
AUTOFIT_MIN_WIDTH = 10


def _css(color: str) -> str:
    return color if color.startswith("#") else f"#{color}"


def _win_path(path: Path) -> str:
    return str(path).replace("/", "\\")


def header_footer_run(color: str, text: str, *, size: str = HF_FONT_SIZE) -> str:
    """Excel 頁首/頁尾片段：顏色 + 兩位字級 + 文字。

    ``&nn`` 後面加空白，避免日期開頭數字被吃進字級碼（``&8`` + ``2026`` → ``&82``）。
    """
    return f"&K{color}&{size} {text}"


def _is_excel_formula(value: Any) -> bool:
    return isinstance(value, str) and value.startswith("=")


def excel_absolute_external_formula(source_path: Path, sheet_name: str, cell_addr: str) -> str:
    """絕對路徑外部連結：``='C:\\folder\\[file.xlsx]Sheet'!A1``。

    報告若寫在 ``run_dir``（與 VBA 中繼檔不同資料夾），相對連結會斷；一律用絕對路徑。
    """
    folder = _win_path(source_path.parent).rstrip("\\")
    filename = source_path.name
    escaped = sheet_name.replace("'", "''")
    return f"='{folder}\\[{filename}]{escaped}'!{cell_addr}"


def build_report(
    config: AppConfig,
    *,
    as_of: date,
    step2_path: Path,
    bbg_values: tuple[tuple[Any, ...], ...],
    bbg_formats: tuple[tuple[str, ...], ...],
    output_path: Path | None = None,
) -> Path:
    """依 config.chart.engine 產出 ``LME每日報價{yyyymmdd}.xlsx`` 到 ``run_dir``。"""
    dest = output_path or config.output_workbook_path(as_of)
    dest.parent.mkdir(parents=True, exist_ok=True)
    logger.info("開始產生報告（engine=%s）：%s", config.chart.engine, dest)
    logger.info("圖表/原始數據來源（vba_dir 中繼檔）：%s", step2_path)
    logger.debug("%s：%s", SOURCE_LABEL, _win_path(step2_path))

    try:
        step2_ready = step2_path.is_file()
    except OSError as exc:
        raise ReportBuildError(f"無法確認步驟二產出檔：{step2_path}（{exc}）") from exc
    if not step2_ready:
        raise ReportBuildError(
            f"步驟二產出檔不存在：{step2_path}。"
            "「遠期走勢圖」與「原始數據」只讀 vba_dir 下的 yyyymmdd.xlsx，中斷而不產生缺資料報告。"
        )

    curve_df = load_forward_curve(step2_path)
    plot_df = slice_plot_window(curve_df, forward_months=config.chart.forward_months)

    if config.chart.engine == "xlsxwriter":
        _build_with_xlsxwriter(
            dest, step2_path, curve_df, plot_df, bbg_values, bbg_formats, config, as_of
        )
    else:
        _build_with_openpyxl(dest, step2_path, plot_df, bbg_values, bbg_formats, config, as_of)

    if not dest.is_file():
        raise ReportBuildError(f"報告寫入後檔案不存在：{dest}")
    logger.info("報告已寫入：%s", dest)
    return dest


def load_forward_curve(step2_path: Path) -> pd.DataFrame:
    """讀取步驟二單表，統一 Prompt 日期；轉換失敗列記警告但不中斷。"""
    try:
        df = pd.read_excel(step2_path)
    except Exception as exc:
        raise ReportBuildError(f"無法讀取 {step2_path}：{exc}") from exc

    if DATE_COL not in df.columns:
        raise ReportBuildError(
            f"{step2_path.name} 缺少欄位 {DATE_COL!r}。實際欄位：{list(df.columns)}"
        )
    missing_series = [code for code in SIX_SERIES if code not in df.columns]
    if missing_series:
        raise ReportBuildError(
            f"{step2_path.name} 缺少品種欄位 {missing_series}。實際欄位：{list(df.columns)}"
        )

    parsed = pd.to_datetime(df[DATE_COL], dayfirst=True, errors="coerce", format="mixed")
    original_na = df[DATE_COL].isna().sum()
    failed = int(parsed.isna().sum() - original_na)
    if failed:
        bad_rows = df.loc[parsed.isna() & df[DATE_COL].notna(), DATE_COL]
        sample = bad_rows.astype(str).head(8).tolist()
        logger.warning(
            "Prompt 日期轉換失敗 %d 列（不中斷，這些列不會進入圖表）。樣本：%s",
            failed,
            sample,
        )
    out = df.copy()
    out[DATE_COL] = parsed
    if out[DATE_COL].notna().sum() == 0:
        raise ReportBuildError(f"{step2_path.name} 的 {DATE_COL} 全部無法解析為日期")
    logger.info("遠期曲線列數=%d，有效日期=%d", len(out), int(out[DATE_COL].notna().sum()))
    return out


def slice_plot_window(df: pd.DataFrame, *, forward_months: int) -> pd.DataFrame:
    """cash_date = min(Prompt)；只保留 cash_date 起 ``forward_months`` 個月內的列。"""
    valid = df.dropna(subset=[DATE_COL])
    cash_date = valid[DATE_COL].min()
    cutoff = cash_date + relativedelta(months=forward_months)
    plot_df = valid[valid[DATE_COL] <= cutoff].sort_values(DATE_COL)
    logger.info(
        "圖表視窗：cash_date=%s cutoff=%s（+%dmo）列數 %d → %d",
        pd.Timestamp(cash_date).date(),
        pd.Timestamp(cutoff).date(),
        forward_months,
        len(df),
        len(plot_df),
    )
    if plot_df.empty:
        raise ReportBuildError("過濾 27 個月視窗後沒有可畫圖的資料")
    return plot_df.reset_index(drop=True)


def _excel_serial_to_datetime(value: Any) -> Any:
    if isinstance(value, (datetime, date)):
        return value
    if isinstance(value, (int, float)) and 20000 < float(value) < 80000:
        return EXCEL_EPOCH + timedelta(days=float(value))
    return value


def _looks_like_date_format(fmt: str | None) -> bool:
    if not fmt:
        return False
    lowered = fmt.lower()
    return any(token in lowered for token in ("y", "d", "m", "年", "月", "日")) and "h" not in lowered[:3]


def _thin_border() -> Border:
    side = Side(style="thin", color=COLOR_LINE)
    return Border(left=side, right=side, top=side, bottom=side)


def _is_number_like(value: Any) -> bool:
    if value is None or isinstance(value, (bool, str, datetime, date)):
        return False
    if isinstance(value, (int, float)):
        return True
    try:
        float(value)
    except (TypeError, ValueError):
        return False
    return True


def _prepare_written_value(raw: Any, fmt: str | None = None, *, is_header: bool = False) -> tuple[Any, str | None]:
    """normalize 後寫入；數值已在讀取時依 Excel 顯示文字解析。回傳 (value, number_format)。"""
    value = normalize_bbg_cell(raw)
    if is_header:
        return value, None
    if _looks_like_date_format(fmt):
        value = _excel_serial_to_datetime(value)
        return value, "YYYY-MM-DD"
    if _is_number_like(value):
        return value, NUMBER_FORMAT_2DP
    return value, None


def _cell_display_text(value: Any, *, number_format: str | None = None) -> str:
    if value is None:
        return ""
    fmt = number_format or ""
    if _is_number_like(value) and ("0.00" in fmt or fmt in {"", "General", "G"}):
        return f"{float(value):,.2f}"
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    return str(value)


def _visual_len(text: str) -> int:
    return sum(2 if ord(ch) > 127 else 1 for ch in text)


def _autofit_width(max_len: int) -> float:
    if max_len <= 0:
        return 8
    return max(max_len + AUTOFIT_PADDING, AUTOFIT_MIN_WIDTH)


def apply_autofit_columns(ws: Worksheet, *, padding: int = AUTOFIT_PADDING) -> None:
    """依內容自動調整欄寬。必須在該 sheet 資料全部寫入後呼叫。"""
    for col_cells in ws.columns:
        max_len = 0
        for cell in col_cells:
            if cell.value is None:
                continue
            text = _cell_display_text(cell.value, number_format=cell.number_format)
            max_len = max(max_len, _visual_len(text))
        letter = get_column_letter(col_cells[0].column)
        if max_len <= 0:
            ws.column_dimensions[letter].width = 8
        else:
            ws.column_dimensions[letter].width = max(max_len + padding, AUTOFIT_MIN_WIDTH)


def apply_center_alignment(ws: Worksheet) -> None:
    """水平＋垂直置中。必須在該 sheet 資料全部寫入後呼叫，只改 alignment。"""
    center = Alignment(horizontal="center", vertical="center")
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = center


class _ColWidthAcc:
    """xlsxwriter 無法回讀儲存格，寫入時累計各欄顯示寬度。"""

    def __init__(self) -> None:
        self._max: dict[int, int] = {}

    def add(self, col: int, value: Any, number_format: str | None = None) -> None:
        text = _cell_display_text(value, number_format=number_format)
        if not text:
            return
        self._max[col] = max(self._max.get(col, 0), _visual_len(text))

    def apply_xlsxwriter(self, ws: Any) -> None:
        for col, max_len in self._max.items():
            ws.set_column(col, col, _autofit_width(max_len))


def _write_bbg_sheet_openpyxl(
    ws: Worksheet,
    values: tuple[tuple[Any, ...], ...],
    formats: tuple[tuple[str, ...], ...],
    *,
    start_row: int = 1,
    header_fill_color: str = COLOR_NAVY,
) -> int:
    ws.sheet_view.showGridLines = True
    header_fill = PatternFill("solid", fgColor=header_fill_color)
    header_font = Font(color=COLOR_WHITE, bold=True)
    last_row = start_row - 1
    for r_idx, row in enumerate(values):
        fmt_row = formats[r_idx] if r_idx < len(formats) else ()
        excel_row = start_row + r_idx
        last_row = excel_row
        for c_idx, raw in enumerate(row):
            fmt = fmt_row[c_idx] if c_idx < len(fmt_row) else None
            value, num_fmt = _prepare_written_value(raw, fmt, is_header=(r_idx == 0))
            cell = ws.cell(row=excel_row, column=c_idx + 1, value=value)
            cell.border = _thin_border()
            if num_fmt:
                cell.number_format = num_fmt
            elif isinstance(value, datetime):
                cell.number_format = "YYYY-MM-DD"
            elif isinstance(value, date):
                cell.number_format = "YYYY-MM-DD"
            if r_idx == 0:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center")
    return last_row


def _copy_raw_sheet(
    source_path: Path,
    dest_ws: Worksheet,
    *,
    start_row: int = 1,
) -> int:
    """把步驟二中繼檔的「值」貼進目標表，不寫公式、不建外部連結。"""
    src_wb = load_workbook(source_path, data_only=True)
    last_row = start_row - 1
    try:
        src_ws = src_wb.active
        for row in src_ws.iter_rows():
            for cell in row:
                target_row = cell.row + start_row - 1
                last_row = max(last_row, target_row)
                value = cell.value
                if _is_excel_formula(value):
                    logger.debug("原始數據略過公式儲存格 %s", cell.coordinate)
                    continue
                is_header = cell.row == 1
                if not is_header:
                    value = value_from_stored_number(value)
                target = dest_ws.cell(row=target_row, column=cell.column, value=value)
                if _is_number_like(value):
                    target.number_format = NUMBER_FORMAT_2DP
                elif cell.number_format:
                    target.number_format = cell.number_format
                if cell.has_style and cell.font and cell.font.bold:
                    target.font = Font(bold=True)
    finally:
        src_wb.close()
    return last_row


def _apply_print_layout_openpyxl(
    ws: Worksheet,
    *,
    header: str,
    tab_color: str | None = None,
    paper_size: int = PAPERSIZE_A4,
    header_color: str = COLOR_NAVY,
) -> None:
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = paper_size
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.horizontalCentered = True
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddHeader.center.text = header_footer_run(header_color, header)
    ws.oddFooter.center.text = PRINT_FOOTER
    ws.page_margins = PageMargins(
        left=0.4, right=0.4, top=0.75, bottom=0.7, header=0.3, footer=0.4
    )
    ws.print_options.horizontalCentered = True
    if tab_color:
        ws.sheet_properties.tabColor = tab_color


def _set_print_area(ws: Worksheet, last_row: int, last_col: int = 18) -> None:
    if last_row < 1:
        return
    ws.print_area = f"A1:{get_column_letter(last_col)}{last_row}"


def _write_print_disclaimer_openpyxl(ws: Worksheet, *, after_row: int, last_col: int = 18) -> int:
    """合併版面最下方：完整兩段免責聲明，自動換行、深灰小字。"""
    row = after_row + 2
    end = get_column_letter(last_col)
    ws.merge_cells(f"A{row}:{end}{row}")
    cell = ws.cell(row=row, column=1, value=PRINT_DISCLAIMER_FULL)
    cell.font = Font(size=DISCLAIMER_FONT_SIZE, color=COLOR_INK)
    cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
    ws.row_dimensions[row].height = DISCLAIMER_ROW_HEIGHT
    return row


def _write_report_banner_openpyxl(ws: Worksheet, as_of: date, *, last_col: int = 18) -> int:
    """頁首：公司名 + 中英雙語報表標題 + 日期（磚紅底、米白字）。"""
    end = get_column_letter(last_col)
    brick = PatternFill("solid", fgColor=COLOR_BRICK)
    accent = PatternFill("solid", fgColor=COLOR_ACCENT)
    center = Alignment(horizontal="center", vertical="center")

    ws.merge_cells(f"A1:{end}1")
    ws["A1"] = f"{COMPANY_ZH}  /  {COMPANY_EN}"
    ws["A1"].font = Font(color=COLOR_CREAM, size=11, bold=True)
    ws["A1"].fill = brick
    ws["A1"].alignment = center
    ws.row_dimensions[1].height = 18

    ws.merge_cells(f"A2:{end}2")
    ws["A2"] = REPORT_TITLE_ZH
    ws["A2"].font = Font(color=COLOR_CREAM, size=18, bold=True)
    ws["A2"].fill = brick
    ws["A2"].alignment = center
    ws.row_dimensions[2].height = 24

    ws.merge_cells(f"A3:{end}3")
    ws["A3"] = REPORT_TITLE_EN
    ws["A3"].font = Font(color=COLOR_CREAM, size=11)
    ws["A3"].fill = brick
    ws["A3"].alignment = center
    ws.row_dimensions[3].height = 16

    ws.merge_cells(f"A4:{end}4")
    ws["A4"] = f"報價日期 / As of  {as_of.strftime('%Y-%m-%d')}"
    ws["A4"].font = Font(color=COLOR_TITLE_DARK, size=11, bold=True)
    ws["A4"].fill = PatternFill("solid", fgColor=COLOR_GRAY)
    ws["A4"].alignment = center
    ws.row_dimensions[4].height = 18

    ws.merge_cells(f"A5:{end}5")
    ws["A5"].fill = accent
    ws.row_dimensions[5].height = 4
    return 6


def _section_heading_openpyxl(
    ws: Worksheet, row: int, zh: str, en: str | None = None, *, last_col: int = 18
) -> int:
    end = get_column_letter(last_col)
    brick = PatternFill("solid", fgColor=COLOR_BRICK)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
    cell = ws.cell(row=row, column=1, value=f"{zh}  /  {en}" if en else zh)
    cell.font = Font(color=COLOR_CREAM, bold=True, size=12)
    cell.fill = brick
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 18
    for col in range(1, last_col + 1):
        ws.cell(row=row, column=col).fill = brick
    ws[f"A{row}"].fill = brick
    ws[f"{end}{row}"].fill = brick
    return row + 1


def _chart_image_anchors(start_row: int) -> list[str]:
    anchors: list[str] = []
    for pair in range(3):
        row = start_row + pair * CHART_ROW_STRIDE
        anchors.extend((f"A{row}", f"J{row}"))
    return anchors


def _place_chart_images(
    ws: Worksheet,
    images: list[Path],
    config: AppConfig,
    *,
    start_row: int,
) -> int:
    anchors = _chart_image_anchors(start_row)
    for img_path, anchor in zip(images, anchors, strict=True):
        img = XLImage(str(img_path))
        img.width = config.chart.image_width
        img.height = config.chart.image_height
        ws.add_image(img, anchor)
    ws.column_dimensions["A"].width = max(ws.column_dimensions["A"].width or 0, 18)
    ws.column_dimensions["J"].width = max(ws.column_dimensions["J"].width or 0, 18)
    return start_row + 3 * CHART_ROW_STRIDE


def _write_print_sheet_openpyxl(
    wb: Workbook,
    *,
    as_of: date,
    step2_path: Path,
    bbg_values: tuple[tuple[Any, ...], ...],
    bbg_formats: tuple[tuple[str, ...], ...],
    config: AppConfig,
    chart_images: list[Path],
) -> None:
    ws = wb.create_sheet(SHEET_PRINT)
    _write_report_banner_openpyxl(ws, as_of)
    header = f"{REPORT_TITLE_ZH} / {REPORT_TITLE_EN}  {as_of.strftime('%Y-%m-%d')}"

    bbg_heading = 6
    _section_heading_openpyxl(ws, bbg_heading, SHEET_BBG, "Bloomberg Snapshot")
    last = _write_bbg_sheet_openpyxl(
        ws, bbg_values, bbg_formats, start_row=bbg_heading + 1, header_fill_color=COLOR_BRICK
    )

    chart_heading = last + 2
    _section_heading_openpyxl(ws, chart_heading, SHEET_CHART, "Forward Curve")
    after_images = _place_chart_images(ws, chart_images, config, start_row=chart_heading + 2)

    raw_heading = after_images + 1
    _section_heading_openpyxl(ws, raw_heading, PRINT_RAW_SECTION_TITLE)
    last_raw = _copy_raw_sheet(step2_path, ws, start_row=raw_heading + 1)

    disc_row = _write_print_disclaimer_openpyxl(ws, after_row=max(last_raw, after_images))
    _apply_print_layout_openpyxl(
        ws,
        header=header,
        tab_color=COLOR_ACCENT,
        paper_size=PAPERSIZE_A3,
        header_color=COLOR_TITLE_DARK,
    )
    _set_print_area(ws, disc_row)
    ws.oddHeader.left.text = header_footer_run(COLOR_TITLE_DARK, COMPANY_ZH)
    ws.oddHeader.right.text = header_footer_run(COLOR_TITLE_DARK, as_of.strftime("%Y-%m-%d"))


def _build_with_openpyxl(
    dest: Path,
    step2_path: Path,
    plot_df: pd.DataFrame,
    bbg_values: tuple[tuple[Any, ...], ...],
    bbg_formats: tuple[tuple[str, ...], ...],
    config: AppConfig,
    as_of: date,
) -> None:
    wb = Workbook()
    header = f"{REPORT_TITLE_ZH} {as_of.strftime('%Y-%m-%d')}"
    ws_bbg = wb.active
    ws_bbg.title = SHEET_BBG
    _write_bbg_sheet_openpyxl(ws_bbg, bbg_values, bbg_formats)
    apply_autofit_columns(ws_bbg)
    apply_center_alignment(ws_bbg)
    _apply_print_layout_openpyxl(ws_bbg, header=f"{header} · {SHEET_BBG}", tab_color=COLOR_NAVY)

    tmp_dir, images = _render_matplotlib_pngs(plot_df, config)
    try:
        ws_chart = wb.create_sheet(SHEET_CHART)
        ws_chart["A1"] = "LME 遠期曲線（cash date → +{} 個月）".format(config.chart.forward_months)
        ws_chart["A1"].font = Font(bold=True, size=14, color=COLOR_NAVY)
        logger.debug("%s：%s", SOURCE_LABEL, _win_path(step2_path))
        _place_chart_images(ws_chart, images, config, start_row=4)
        apply_autofit_columns(ws_chart)
        _apply_print_layout_openpyxl(
            ws_chart, header=f"{header} · {SHEET_CHART}", tab_color=COLOR_GOLD
        )

        ws_raw = wb.create_sheet(SHEET_RAW)
        _copy_raw_sheet(step2_path, ws_raw)
        apply_autofit_columns(ws_raw)
        _apply_print_layout_openpyxl(
            ws_raw, header=f"{header} · {SHEET_RAW}", tab_color=COLOR_MUTED
        )

        _write_print_sheet_openpyxl(
            wb,
            as_of=as_of,
            step2_path=step2_path,
            bbg_values=bbg_values,
            bbg_formats=bbg_formats,
            config=config,
            chart_images=images,
        )

        try:
            wb.save(dest)
        except Exception as exc:
            raise ReportBuildError(f"儲存報告失敗（{dest}）：{exc}") from exc
    finally:
        wb.close()
        shutil.rmtree(tmp_dir, ignore_errors=True)


def _configure_cjk_matplotlib() -> None:
    """讓圖表標題的中文在 Windows（微軟正黑體 / 雅黑）能正常顯示。"""
    import matplotlib
    from matplotlib import font_manager

    matplotlib.rcParams["axes.unicode_minus"] = False
    candidates = [
        Path(r"C:\Windows\Fonts\msjh.ttc"),
        Path(r"C:\Windows\Fonts\msjhbd.ttc"),
        Path(r"C:\Windows\Fonts\msyh.ttc"),
        Path(r"C:\Windows\Fonts\simhei.ttf"),
        Path("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"),
        Path("/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"),
        Path("/usr/share/fonts/opentype/noto/NotoSansCJKtc-Regular.otf"),
    ]
    for font_path in candidates:
        if font_path.is_file():
            try:
                font_manager.fontManager.addfont(str(font_path))
                name = font_manager.FontProperties(fname=str(font_path)).get_name()
                matplotlib.rcParams["font.sans-serif"] = [name, "DejaVu Sans"]
                matplotlib.rcParams["font.family"] = "sans-serif"
                logger.info("matplotlib 使用中文字型：%s", font_path)
                return
            except Exception as exc:
                logger.debug("載入字型失敗 %s：%s", font_path, exc)
    logger.warning(
        "找不到中文字型，遠期曲線標題可能出現方塊。Windows 請確認已安裝微軟正黑體。"
    )


def _render_matplotlib_pngs(plot_df: pd.DataFrame, config: AppConfig) -> tuple[Path, list[Path]]:
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    _configure_cjk_matplotlib()
    tmp_dir = Path(tempfile.mkdtemp(prefix="lme_charts_"))
    images: list[Path] = []
    for code, title in SIX_SERIES.items():
        series_df = plot_df.dropna(subset=[code])[[DATE_COL, code]]
        if series_df.empty:
            logger.warning("品種 %s 在圖表視窗內沒有非空資料，仍會放空圖說明", code)
        fig, ax = plt.subplots(figsize=(6.4, 3.6), dpi=120)
        fig.patch.set_facecolor("white")
        ax.set_facecolor("white")
        if not series_df.empty:
            y = series_df[code].map(lambda v: value_from_stored_number(v) if pd.notna(v) else v)
            ax.plot(
                series_df[DATE_COL],
                y,
                color=_css(COLOR_NAVY),
                linewidth=1.7,
            )
        ax.set_title(f"{title} 遠期曲線", color=_css(COLOR_NAVY), fontweight="bold")
        ax.set_xlabel("到期日 Prompt", color=_css(COLOR_MUTED))
        ax.set_ylabel("價格", color=_css(COLOR_MUTED))
        ax.tick_params(colors=_css(COLOR_INK))
        ax.grid(True, color="#E5E7EB", alpha=0.95)
        for spine in ax.spines.values():
            spine.set_color(_css(COLOR_GOLD))
            spine.set_linewidth(0.8)
        fig.autofmt_xdate()
        fig.tight_layout()
        img_path = tmp_dir / f"{code}.png"
        fig.savefig(img_path, dpi=120, facecolor=fig.get_facecolor())
        plt.close(fig)
        images.append(img_path)
    return tmp_dir, images


def _build_with_xlsxwriter(
    dest: Path,
    step2_path: Path,
    curve_df: pd.DataFrame,
    plot_df: pd.DataFrame,
    bbg_values: tuple[tuple[Any, ...], ...],
    bbg_formats: tuple[tuple[str, ...], ...],
    config: AppConfig,
    as_of: date,
) -> None:
    import xlsxwriter

    tmp_dir: Path | None = None
    try:
        workbook = xlsxwriter.Workbook(str(dest))
    except Exception as exc:
        raise ReportBuildError(f"無法建立 xlsxwriter 工作簿（{dest}）：{exc}") from exc

    try:
        header = f"{REPORT_TITLE_ZH} {as_of.strftime('%Y-%m-%d')}"
        _write_bbg_sheet_xlsxwriter(workbook, bbg_values, bbg_formats, header=header)
        _write_xlsxwriter_charts(workbook, plot_df, config, header=header, step2_path=step2_path)
        _write_raw_sheet_xlsxwriter(workbook, step2_path, curve_df, header=header)
        tmp_dir, images = _render_matplotlib_pngs(plot_df, config)
        _write_print_sheet_xlsxwriter(
            workbook,
            as_of=as_of,
            step2_path=step2_path,
            curve_df=curve_df,
            bbg_values=bbg_values,
            bbg_formats=bbg_formats,
            config=config,
            chart_images=images,
            header=header,
        )
        workbook.close()
    except ReportBuildError:
        try:
            workbook.close()
        except Exception:
            pass
        raise
    except Exception as exc:
        try:
            workbook.close()
        except Exception:
            pass
        raise ReportBuildError(f"xlsxwriter 產生報告失敗：{exc}") from exc
    finally:
        if tmp_dir:
            shutil.rmtree(tmp_dir, ignore_errors=True)


def _apply_print_layout_xlsxwriter(
    ws: Any, *, header: str, paper: int = 9, header_color: str = COLOR_NAVY
) -> None:
    ws.set_landscape()
    ws.set_paper(paper)
    ws.fit_to_pages(1, 0)
    ws.set_header(f"&C{header_footer_run(header_color, header)}")
    ws.set_footer(f"&C{PRINT_FOOTER}")
    ws.center_horizontally()
    ws.set_margins(left=0.4, right=0.4, top=0.75, bottom=0.7)


def _write_bbg_rows_xlsxwriter(
    ws: Any,
    workbook: Any,
    values: tuple[tuple[Any, ...], ...],
    formats: tuple[tuple[str, ...], ...],
    *,
    start_row: int = 0,
    header_fill_color: str = COLOR_NAVY,
    center_align: bool = False,
) -> int:
    align = {"align": "center", "valign": "vcenter"} if center_align else {}
    header_fmt = workbook.add_format(
        {
            "bold": True,
            "bg_color": _css(header_fill_color),
            "font_color": _css(COLOR_WHITE),
            **align,
        }
    )
    date_fmt = workbook.add_format({"num_format": "yyyy-mm-dd", **align})
    price_fmt = workbook.add_format({"num_format": NUMBER_FORMAT_2DP, **align})
    default_fmt = workbook.add_format(align) if center_align else None
    widths = _ColWidthAcc()
    last = start_row - 1
    for r_idx, row in enumerate(values):
        fmt_row = formats[r_idx] if r_idx < len(formats) else ()
        excel_row = start_row + r_idx
        last = excel_row
        for c_idx, raw in enumerate(row):
            fmt = fmt_row[c_idx] if c_idx < len(fmt_row) else None
            value, num_fmt = _prepare_written_value(raw, fmt, is_header=(r_idx == 0))
            cell_fmt = header_fmt if r_idx == 0 else None
            if num_fmt == "YYYY-MM-DD" and r_idx != 0:
                if isinstance(value, date) and not isinstance(value, datetime):
                    value = datetime(value.year, value.month, value.day)
                cell_fmt = date_fmt
            elif num_fmt == NUMBER_FORMAT_2DP:
                cell_fmt = price_fmt
            if isinstance(value, date) and not isinstance(value, datetime) and r_idx != 0:
                value = datetime(value.year, value.month, value.day)
            if cell_fmt is None:
                cell_fmt = default_fmt
            ws.write(excel_row, c_idx, value, cell_fmt)
            widths.add(c_idx, value, number_format=num_fmt)
    widths.apply_xlsxwriter(ws)
    return last


def _write_bbg_sheet_xlsxwriter(
    workbook: Any,
    values: tuple[tuple[Any, ...], ...],
    formats: tuple[tuple[str, ...], ...],
    *,
    header: str,
) -> None:
    ws = workbook.add_worksheet(SHEET_BBG)
    _apply_print_layout_xlsxwriter(ws, header=f"{header} · {SHEET_BBG}")
    ws.set_tab_color(_css(COLOR_NAVY))
    _write_bbg_rows_xlsxwriter(
        ws, workbook, values, formats, start_row=0, center_align=True
    )


def _write_xlsxwriter_charts(
    workbook: Any,
    plot_df: pd.DataFrame,
    config: AppConfig,
    *,
    header: str,
    step2_path: Path,
) -> None:
    data_ws = workbook.add_worksheet(SHEET_CHART_DATA)
    chart_ws = workbook.add_worksheet(SHEET_CHART)
    data_ws.hide()
    title_fmt = workbook.add_format(
        {"bold": True, "font_size": 14, "font_color": _css(COLOR_NAVY)}
    )
    chart_ws.write(0, 0, f"LME 遠期曲線（cash date → +{config.chart.forward_months} 個月）", title_fmt)
    logger.debug("%s：%s", SOURCE_LABEL, _win_path(step2_path))
    _apply_print_layout_xlsxwriter(chart_ws, header=f"{header} · {SHEET_CHART}")
    chart_ws.set_tab_color(_css(COLOR_GOLD))

    date_fmt = workbook.add_format({"num_format": "yyyy-mm-dd"})
    price_fmt = workbook.add_format({"num_format": NUMBER_FORMAT_2DP})
    positions = [
        ("A4", "J4"),
        ("A23", "J23"),
        ("A42", "J42"),
    ]
    data_widths = _ColWidthAcc()
    col = 0
    for idx, (code, title) in enumerate(SIX_SERIES.items()):
        series_df = plot_df.dropna(subset=[code])[[DATE_COL, code]].reset_index(drop=True)
        data_ws.write(0, col, f"{code}_date")
        data_ws.write(0, col + 1, f"{code}_px")
        data_widths.add(col, f"{code}_date")
        data_widths.add(col + 1, f"{code}_px")
        for r, rec in series_df.iterrows():
            ts = rec[DATE_COL]
            if isinstance(ts, pd.Timestamp):
                ts = ts.to_pydatetime()
            px = value_from_stored_number(rec[code])
            data_ws.write_datetime(r + 1, col, ts, date_fmt)
            data_ws.write_number(r + 1, col + 1, float(px), price_fmt)
            data_widths.add(col, ts, number_format="yyyy-mm-dd")
            data_widths.add(col + 1, px, number_format=NUMBER_FORMAT_2DP)

        chart = workbook.add_chart({"type": "line"})
        n = len(series_df)
        if n:
            chart.add_series(
                {
                    "name": title,
                    "categories": [SHEET_CHART_DATA, 1, col, n, col],
                    "values": [SHEET_CHART_DATA, 1, col + 1, n, col + 1],
                    "line": {"color": _css(COLOR_NAVY), "width": 1.5},
                }
            )
        chart.set_title({"name": f"{title} 遠期曲線"})
        chart.set_x_axis({"name": "到期日 Prompt", "num_format": "yyyy-mm-dd"})
        chart.set_y_axis({"name": "價格"})
        chart.set_size({"width": config.chart.image_width, "height": config.chart.image_height})
        chart.set_legend({"none": True})
        row_pair, col_slot = divmod(idx, 2)
        anchor = positions[row_pair][col_slot]
        chart_ws.insert_chart(anchor, chart)
        col += 2
    data_widths.apply_xlsxwriter(data_ws)
    chart_widths = _ColWidthAcc()
    chart_widths.add(0, f"LME 遠期曲線（cash date → +{config.chart.forward_months} 個月）")
    chart_widths.apply_xlsxwriter(chart_ws)


def _write_raw_sheet_xlsxwriter(
    workbook: Any,
    step2_path: Path,
    curve_df: pd.DataFrame,
    *,
    header: str,
    worksheet: Any | None = None,
    start_row: int = 0,
) -> int:
    ws = worksheet or workbook.add_worksheet(SHEET_RAW)
    if worksheet is None:
        _apply_print_layout_xlsxwriter(ws, header=f"{header} · {SHEET_RAW}")
        ws.set_tab_color(_css(COLOR_MUTED))
    header_fmt = workbook.add_format({"bold": True})
    date_fmt = workbook.add_format({"num_format": "yyyy-mm-dd"})
    price_fmt = workbook.add_format({"num_format": NUMBER_FORMAT_2DP})
    widths = _ColWidthAcc()
    own_sheet = worksheet is None
    last = start_row - 1
    try:
        src_wb = load_workbook(step2_path, data_only=True)
        src_ws = src_wb.active
        for row in src_ws.iter_rows():
            for cell in row:
                r, c = cell.row - 1 + start_row, cell.column - 1
                last = max(last, r)
                value = cell.value
                if _is_excel_formula(value):
                    logger.debug("原始數據略過公式儲存格 %s", cell.coordinate)
                    continue
                is_header = cell.row == 1
                if not is_header:
                    value = value_from_stored_number(value)
                fmt = header_fmt if is_header else None
                num_fmt: str | None = None
                if isinstance(value, datetime):
                    ws.write_datetime(r, c, value, date_fmt)
                    num_fmt = "yyyy-mm-dd"
                elif isinstance(value, date):
                    ws.write_datetime(r, c, datetime(value.year, value.month, value.day), date_fmt)
                    num_fmt = "yyyy-mm-dd"
                elif _is_number_like(value):
                    ws.write_number(r, c, float(value), price_fmt)
                    num_fmt = NUMBER_FORMAT_2DP
                else:
                    ws.write(r, c, value, fmt)
                widths.add(c, value, number_format=num_fmt)
        src_wb.close()
        if own_sheet:
            widths.apply_xlsxwriter(ws)
        return last
    except Exception as exc:
        logger.warning("xlsxwriter 逐格複製原始數據失敗，改用 pandas：%s", exc)

    for c_idx, col_name in enumerate(curve_df.columns):
        ws.write(start_row, c_idx, col_name, header_fmt)
        widths.add(c_idx, col_name)
        last = start_row
    for r_idx, rec in curve_df.iterrows():
        excel_row = int(r_idx) + 1 + start_row
        last = max(last, excel_row)
        for c_idx, col_name in enumerate(curve_df.columns):
            value = rec[col_name]
            if pd.isna(value):
                continue
            if col_name == DATE_COL and isinstance(value, pd.Timestamp):
                ws.write_datetime(excel_row, c_idx, value.to_pydatetime(), date_fmt)
                widths.add(c_idx, value.to_pydatetime(), number_format="yyyy-mm-dd")
            else:
                value = value_from_stored_number(value)
                if _is_number_like(value):
                    ws.write_number(excel_row, c_idx, float(value), price_fmt)
                    widths.add(c_idx, value, number_format=NUMBER_FORMAT_2DP)
                else:
                    ws.write(excel_row, c_idx, value)
                    widths.add(c_idx, value)
    if own_sheet:
        widths.apply_xlsxwriter(ws)
    return last


def _write_print_sheet_xlsxwriter(
    workbook: Any,
    *,
    as_of: date,
    step2_path: Path,
    curve_df: pd.DataFrame,
    bbg_values: tuple[tuple[Any, ...], ...],
    bbg_formats: tuple[tuple[str, ...], ...],
    config: AppConfig,
    chart_images: list[Path],
    header: str,
) -> None:
    ws = workbook.add_worksheet(SHEET_PRINT)
    brick = workbook.add_format(
        {
            "bold": True,
            "font_color": _css(COLOR_CREAM),
            "bg_color": _css(COLOR_BRICK),
            "align": "center",
            "valign": "vcenter",
        }
    )
    sub = workbook.add_format(
        {
            "font_color": _css(COLOR_CREAM),
            "bg_color": _css(COLOR_BRICK),
            "align": "center",
        }
    )
    date_fmt = workbook.add_format(
        {
            "bold": True,
            "font_color": _css(COLOR_TITLE_DARK),
            "bg_color": _css(COLOR_GRAY),
            "align": "center",
        }
    )
    accent = workbook.add_format({"bg_color": _css(COLOR_ACCENT)})
    section = workbook.add_format(
        {
            "bold": True,
            "font_color": _css(COLOR_CREAM),
            "bg_color": _css(COLOR_BRICK),
            "align": "left",
        }
    )
    last_col = 17
    ws.merge_range(0, 0, 0, last_col, f"{COMPANY_ZH}  /  {COMPANY_EN}", brick)
    ws.merge_range(1, 0, 1, last_col, REPORT_TITLE_ZH, brick)
    ws.merge_range(2, 0, 2, last_col, REPORT_TITLE_EN, sub)
    ws.merge_range(3, 0, 3, last_col, f"報價日期 / As of  {as_of.strftime('%Y-%m-%d')}", date_fmt)
    ws.merge_range(4, 0, 4, last_col, "", accent)
    ws.set_row(1, 22)

    ws.merge_range(5, 0, 5, last_col, f"{SHEET_BBG}  /  Bloomberg Snapshot", section)
    last = _write_bbg_rows_xlsxwriter(
        ws, workbook, bbg_values, bbg_formats, start_row=6, header_fill_color=COLOR_BRICK
    )

    chart_heading = last + 2
    ws.merge_range(chart_heading, 0, chart_heading, last_col, f"{SHEET_CHART}  /  Forward Curve", section)
    img_start = chart_heading + 2
    native_w = 6.4 * 120
    native_h = 3.6 * 120
    x_scale = config.chart.image_width / native_w
    y_scale = config.chart.image_height / native_h
    anchors = _chart_image_anchors(img_start + 1)
    for img_path, anchor in zip(chart_images, anchors, strict=True):
        ws.insert_image(anchor, str(img_path), {"x_scale": x_scale, "y_scale": y_scale})
    after_images = img_start + 3 * CHART_ROW_STRIDE

    raw_heading = after_images
    ws.merge_range(raw_heading, 0, raw_heading, last_col, PRINT_RAW_SECTION_TITLE, section)
    last_raw = _write_raw_sheet_xlsxwriter(
        workbook,
        step2_path,
        curve_df,
        header=header,
        worksheet=ws,
        start_row=raw_heading + 1,
    )
    disc_row = max(last_raw, after_images) + 2
    disc_fmt = workbook.add_format(
        {
            "font_size": DISCLAIMER_FONT_SIZE,
            "font_color": _css(COLOR_INK),
            "text_wrap": True,
            "valign": "top",
            "align": "left",
        }
    )
    ws.merge_range(disc_row, 0, disc_row, last_col, PRINT_DISCLAIMER_FULL, disc_fmt)
    ws.set_row(disc_row, DISCLAIMER_ROW_HEIGHT)
    _apply_print_layout_xlsxwriter(
        ws, header=header, paper=PAPERSIZE_A3, header_color=COLOR_TITLE_DARK
    )
    ws.set_header(
        f"&L{header_footer_run(COLOR_TITLE_DARK, COMPANY_ZH)}"
        f"&C{header_footer_run(COLOR_TITLE_DARK, header)}"
        f"&R{header_footer_run(COLOR_TITLE_DARK, as_of.strftime('%Y-%m-%d'))}"
    )
    ws.set_footer(f"&C{PRINT_FOOTER}")
    ws.set_tab_color(_css(COLOR_ACCENT))
    ws.print_area(0, 0, disc_row, last_col)


def dataframe_preview_rows(df: pd.DataFrame, limit: int = 5) -> list[list[Any]]:
    """測試輔助：把 DataFrame 前幾列轉成純 Python。"""
    rows = []
    for row in dataframe_to_rows(df.head(limit), index=False, header=True):
        rows.append(list(row))
    return rows
