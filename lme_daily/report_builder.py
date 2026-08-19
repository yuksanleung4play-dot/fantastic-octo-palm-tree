"""產生最終報告 Excel：BBG快照 / 遠期走勢圖 / 原始數據。

圖表引擎由 ``chart.engine`` 切換：
- ``matplotlib``：靜態 PNG 嵌入 openpyxl（預設）
- ``xlsxwriter``：Excel 原生可互動折線圖
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
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.worksheet import Worksheet

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
SHEET_CHART_DATA = "_chart_data"

EXCEL_EPOCH = datetime(1899, 12, 30)


def build_report(
    config: AppConfig,
    *,
    as_of: date,
    step2_path: Path,
    bbg_values: tuple[tuple[Any, ...], ...],
    bbg_formats: tuple[tuple[str, ...], ...],
    output_path: Path | None = None,
) -> Path:
    """依 config.chart.engine 產出 ``LME每日報價{yyyymmdd}.xlsx``。"""
    dest = output_path or config.output_workbook_path(as_of)
    dest.parent.mkdir(parents=True, exist_ok=True)
    logger.info("開始產生報告（engine=%s）：%s", config.chart.engine, dest)

    if not step2_path.is_file():
        raise ReportBuildError(f"步驟二產出檔不存在：{step2_path}")

    curve_df = load_forward_curve(step2_path)
    plot_df = slice_plot_window(curve_df, forward_months=config.chart.forward_months)

    if config.chart.engine == "xlsxwriter":
        _build_with_xlsxwriter(dest, step2_path, curve_df, plot_df, bbg_values, bbg_formats, config)
    else:
        _build_with_openpyxl(dest, step2_path, plot_df, bbg_values, bbg_formats, config)

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


def _write_bbg_sheet_openpyxl(
    ws: Worksheet,
    values: tuple[tuple[Any, ...], ...],
    formats: tuple[tuple[str, ...], ...],
) -> None:
    ws.sheet_view.showGridLines = True
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    for r_idx, row in enumerate(values):
        fmt_row = formats[r_idx] if r_idx < len(formats) else ()
        for c_idx, raw in enumerate(row):
            fmt = fmt_row[c_idx] if c_idx < len(fmt_row) else None
            value = raw
            if _looks_like_date_format(fmt):
                value = _excel_serial_to_datetime(raw)
            cell = ws.cell(row=r_idx + 1, column=c_idx + 1, value=value)
            if fmt and fmt not in {"General", "G"}:
                cell.number_format = fmt
            elif isinstance(value, datetime):
                cell.number_format = "YYYY-MM-DD"
            elif isinstance(value, date):
                cell.number_format = "YYYY-MM-DD"
            if r_idx == 0:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center")
    for col in ws.columns:
        letter = col[0].column_letter
        ws.column_dimensions[letter].width = 14


def _copy_raw_sheet(source_path: Path, dest_ws: Worksheet) -> None:
    """把步驟二整表原封不動複製到「原始數據」（含欄名與 number_format）。"""
    src_wb = load_workbook(source_path, data_only=True)
    try:
        src_ws = src_wb.active
        for row in src_ws.iter_rows():
            for cell in row:
                target = dest_ws.cell(row=cell.row, column=cell.column, value=cell.value)
                if cell.number_format:
                    target.number_format = cell.number_format
                if cell.has_style and cell.font and cell.font.bold:
                    target.font = Font(bold=True)
        for col_letter, dim in src_ws.column_dimensions.items():
            if dim.width:
                dest_ws.column_dimensions[col_letter].width = dim.width
    finally:
        src_wb.close()


def _build_with_openpyxl(
    dest: Path,
    step2_path: Path,
    plot_df: pd.DataFrame,
    bbg_values: tuple[tuple[Any, ...], ...],
    bbg_formats: tuple[tuple[str, ...], ...],
    config: AppConfig,
) -> None:
    wb = Workbook()
    ws_bbg = wb.active
    ws_bbg.title = SHEET_BBG
    _write_bbg_sheet_openpyxl(ws_bbg, bbg_values, bbg_formats)

    ws_chart = wb.create_sheet(SHEET_CHART)
    _embed_matplotlib_charts(ws_chart, plot_df, config)
    chart_tmpdir = getattr(ws_chart, "_lme_chart_tmpdir", None)

    ws_raw = wb.create_sheet(SHEET_RAW)
    _copy_raw_sheet(step2_path, ws_raw)

    try:
        wb.save(dest)
    except Exception as exc:
        raise ReportBuildError(f"儲存報告失敗（{dest}）：{exc}") from exc
    finally:
        wb.close()
        if chart_tmpdir:
            shutil.rmtree(chart_tmpdir, ignore_errors=True)


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


def _embed_matplotlib_charts(ws: Worksheet, plot_df: pd.DataFrame, config: AppConfig) -> None:
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    _configure_cjk_matplotlib()

    ws["A1"] = "LME 遠期曲線（cash date → +{} 個月）".format(config.chart.forward_months)
    ws["A1"].font = Font(bold=True, size=14, color="1F4E79")

    tmp_dir = Path(tempfile.mkdtemp(prefix="lme_charts_"))
    images: list[Path] = []
    try:
        for code, title in SIX_SERIES.items():
            series_df = plot_df.dropna(subset=[code])[[DATE_COL, code]]
            if series_df.empty:
                logger.warning("品種 %s 在圖表視窗內沒有非空資料，仍會放空圖說明", code)
            fig, ax = plt.subplots(figsize=(6.4, 3.6), dpi=120)
            if not series_df.empty:
                ax.plot(
                    series_df[DATE_COL],
                    series_df[code],
                    color="#1F4E79",
                    linewidth=1.6,
                )
            ax.set_title(f"{title} 遠期曲線")
            ax.set_xlabel("到期日 Prompt")
            ax.set_ylabel("價格")
            ax.grid(True, alpha=0.3)
            fig.autofmt_xdate()
            fig.tight_layout()
            img_path = tmp_dir / f"{code}.png"
            fig.savefig(img_path, dpi=120)
            plt.close(fig)
            images.append(img_path)

        # 2 欄 × 3 列
        anchors = ["A3", "J3", "A22", "J22", "A41", "J41"]
        for img_path, anchor in zip(images, anchors, strict=True):
            img = XLImage(str(img_path))
            img.width = config.chart.image_width
            img.height = config.chart.image_height
            ws.add_image(img, anchor)
        ws.column_dimensions["A"].width = 18
    finally:
        # openpyxl 在 save 時才讀圖檔，故暫存目錄需活到 workbook.save 之後。
        # 把路徑記在 worksheet 上，save 後由呼叫端無法輕易清；改為讓 Image 讀入後保留檔案，
        # 呼叫端在 workbook.save 後刪除。此處把 tmp_dir 掛到 sheet 註解供除錯。
        ws.sheet_properties.tabColor = "1F4E79"
        ws._lme_chart_tmpdir = tmp_dir  # type: ignore[attr-defined]


def _build_with_xlsxwriter(
    dest: Path,
    step2_path: Path,
    curve_df: pd.DataFrame,
    plot_df: pd.DataFrame,
    bbg_values: tuple[tuple[Any, ...], ...],
    bbg_formats: tuple[tuple[str, ...], ...],
    config: AppConfig,
) -> None:
    import xlsxwriter

    try:
        workbook = xlsxwriter.Workbook(str(dest))
    except Exception as exc:
        raise ReportBuildError(f"無法建立 xlsxwriter 工作簿（{dest}）：{exc}") from exc

    try:
        _write_bbg_sheet_xlsxwriter(workbook, bbg_values, bbg_formats)
        _write_xlsxwriter_charts(workbook, plot_df, config)
        _write_raw_sheet_xlsxwriter(workbook, step2_path, curve_df)
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


def _write_bbg_sheet_xlsxwriter(
    workbook: Any,
    values: tuple[tuple[Any, ...], ...],
    formats: tuple[tuple[str, ...], ...],
) -> None:
    ws = workbook.add_worksheet(SHEET_BBG)
    header_fmt = workbook.add_format({"bold": True, "bg_color": "#1F4E79", "font_color": "#FFFFFF"})
    date_fmt = workbook.add_format({"num_format": "yyyy-mm-dd"})
    for r_idx, row in enumerate(values):
        fmt_row = formats[r_idx] if r_idx < len(formats) else ()
        for c_idx, raw in enumerate(row):
            fmt = fmt_row[c_idx] if c_idx < len(fmt_row) else None
            value = raw
            cell_fmt = header_fmt if r_idx == 0 else None
            if _looks_like_date_format(fmt):
                value = _excel_serial_to_datetime(raw)
                cell_fmt = date_fmt if r_idx != 0 else header_fmt
            if isinstance(value, date) and not isinstance(value, datetime):
                value = datetime(value.year, value.month, value.day)
            ws.write(r_idx, c_idx, value, cell_fmt)
    ws.set_column(0, 12, 14)


def _write_xlsxwriter_charts(workbook: Any, plot_df: pd.DataFrame, config: AppConfig) -> None:
    data_ws = workbook.add_worksheet(SHEET_CHART_DATA)
    chart_ws = workbook.add_worksheet(SHEET_CHART)
    data_ws.hide()
    chart_ws.write(0, 0, f"LME 遠期曲線（cash date → +{config.chart.forward_months} 個月）")

    date_fmt = workbook.add_format({"num_format": "yyyy-mm-dd"})
    positions = [
        ("A3", "J3"),
        ("A22", "J22"),
        ("A41", "J41"),
    ]
    col = 0
    for idx, (code, title) in enumerate(SIX_SERIES.items()):
        series_df = plot_df.dropna(subset=[code])[[DATE_COL, code]].reset_index(drop=True)
        data_ws.write(0, col, f"{code}_date")
        data_ws.write(0, col + 1, f"{code}_px")
        for r, rec in series_df.iterrows():
            ts = rec[DATE_COL]
            if isinstance(ts, pd.Timestamp):
                ts = ts.to_pydatetime()
            data_ws.write_datetime(r + 1, col, ts, date_fmt)
            data_ws.write_number(r + 1, col + 1, float(rec[code]))

        chart = workbook.add_chart({"type": "line"})
        n = len(series_df)
        if n:
            chart.add_series(
                {
                    "name": title,
                    "categories": [SHEET_CHART_DATA, 1, col, n, col],
                    "values": [SHEET_CHART_DATA, 1, col + 1, n, col + 1],
                    "line": {"color": "#1F4E79", "width": 1.5},
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


def _write_raw_sheet_xlsxwriter(workbook: Any, step2_path: Path, curve_df: pd.DataFrame) -> None:
    ws = workbook.add_worksheet(SHEET_RAW)
    header_fmt = workbook.add_format({"bold": True})
    date_fmt = workbook.add_format({"num_format": "yyyy-mm-dd"})
    # 優先從原始 xlsx 逐格複製；失敗則退回 pandas DataFrame
    try:
        src_wb = load_workbook(step2_path, data_only=True)
        src_ws = src_wb.active
        for row in src_ws.iter_rows():
            for cell in row:
                r, c = cell.row - 1, cell.column - 1
                value = cell.value
                fmt = header_fmt if cell.row == 1 else None
                if isinstance(value, datetime):
                    ws.write_datetime(r, c, value, date_fmt)
                elif isinstance(value, date):
                    ws.write_datetime(r, c, datetime(value.year, value.month, value.day), date_fmt)
                else:
                    ws.write(r, c, value, fmt)
        src_wb.close()
        return
    except Exception as exc:
        logger.warning("xlsxwriter 逐格複製原始數據失敗，改用 pandas：%s", exc)

    for c_idx, col_name in enumerate(curve_df.columns):
        ws.write(0, c_idx, col_name, header_fmt)
    for r_idx, rec in curve_df.iterrows():
        for c_idx, col_name in enumerate(curve_df.columns):
            value = rec[col_name]
            if pd.isna(value):
                continue
            if col_name == DATE_COL and isinstance(value, pd.Timestamp):
                ws.write_datetime(int(r_idx) + 1, c_idx, value.to_pydatetime(), date_fmt)
            else:
                ws.write(int(r_idx) + 1, c_idx, value)


def dataframe_preview_rows(df: pd.DataFrame, limit: int = 5) -> list[list[Any]]:
    """測試輔助：把 DataFrame 前幾列轉成純 Python。"""
    rows = []
    for row in dataframe_to_rows(df.head(limit), index=False, header=True):
        rows.append(list(row))
    return rows
