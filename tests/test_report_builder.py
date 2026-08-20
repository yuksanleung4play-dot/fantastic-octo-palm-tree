"""報告產生（matplotlib / xlsxwriter），不需 Excel COM。"""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import load_workbook

from lme_daily.config import load_config
from lme_daily.exceptions import ReportBuildError
from lme_daily.report_builder import (
    DATE_COL,
    SHEET_BBG,
    SHEET_CHART,
    SHEET_PRINT,
    SHEET_RAW,
    SIX_SERIES,
    build_report,
    excel_absolute_external_formula,
    load_forward_curve,
    slice_plot_window,
)


def _write_min_config(tmp_path: Path, engine: str) -> Path:
    import yaml

    work = tmp_path / "work"
    work.mkdir(exist_ok=True)
    payload = {
        "paths": {
            "working_dir": str(work),
            "ref_workbook_name": "ref.xlsm",
            "bbg_workbook_name": "bbg.xlsx",
            "output_prefix": "LME每日報價",
            "output_dir": "",
        },
        "vba": {"macro_name": "RunDailyLME", "use_param_injection": True},
        "bloomberg": {
            "copy_range": "B3:I10",
            "bbg_sheet_name": "Promt date",
            "refresh_wait_seconds": 1,
        },
        "chart": {"forward_months": 27, "engine": engine},
        "logging": {"level": "INFO", "file": ""},
    }
    path = tmp_path / "config.yaml"
    path.write_text(yaml.safe_dump(payload, allow_unicode=True), encoding="utf-8")
    return path


def _winish(value: object) -> str:
    return str(value).replace("/", "\\")


def _make_step2(path: Path) -> None:
    cash = datetime(2026, 8, 19)
    rows = []
    # 混雜日期字串 + datetime；NI/SN 在較長天期留空
    for i in range(0, 40):
        d = cash + pd.DateOffset(months=i) if i % 3 == 0 else cash + pd.Timedelta(days=i * 20)
        prompt: object
        if i % 5 == 0:
            prompt = d.strftime("%d-%b-%y")
        elif i % 5 == 1:
            prompt = d.strftime("%d/%b/%y")
        else:
            prompt = d
        row = {
            DATE_COL: prompt,
            "CA": 9000 + i,
            "AH": 2500 + i,
            "ZS": 2800 + i,
            "NI": 16000 + i if i < 20 else None,
            "SN": 30000 + i if i < 18 else None,
            "PB": 2000 + i,
        }
        rows.append(row)
    # 一列無法解析的日期：應記警告但不中斷
    rows.append(
        {
            DATE_COL: "not-a-date",
            "CA": 1,
            "AH": 1,
            "ZS": 1,
            "NI": 1,
            "SN": 1,
            "PB": 1,
        }
    )
    pd.DataFrame(rows).to_excel(path, index=False)


def _bbg_sample():
    values = (
        ("Metal", "Last", "Bid", "Ask", "High", "Low", "Settle", "AsOf"),
        ("CA", 9001.5, 9000.0, 9002.0, 9050.0, 8980.0, 9001.0, 45953),
    )
    formats = (
        ("General",) * 8,
        ("General", "0.0", "0.0", "0.0", "0.0", "0.0", "0.0", "YYYY-MM-DD"),
    )
    return values, formats


@pytest.mark.parametrize("engine", ["matplotlib", "xlsxwriter"])
def test_build_report_three_sheets(tmp_path: Path, engine: str):
    cfg_path = _write_min_config(tmp_path, engine)
    config = load_config(cfg_path)
    as_of = date(2026, 8, 19)
    vba_dir = config.vba_dir(as_of)
    vba_dir.mkdir(parents=True, exist_ok=True)
    step2 = vba_dir / "20260819.xlsx"
    _make_step2(step2)
    values, formats = _bbg_sample()

    dest = build_report(
        config,
        as_of=as_of,
        step2_path=step2,
        bbg_values=values,
        bbg_formats=formats,
    )
    assert dest.name == "LME每日報價20260819.xlsx"
    assert dest.parent == config.run_dir(as_of)
    assert dest.is_file()

    wb = load_workbook(dest)
    assert SHEET_BBG in wb.sheetnames
    assert SHEET_CHART in wb.sheetnames
    assert SHEET_RAW in wb.sheetnames
    assert SHEET_PRINT in wb.sheetnames
    assert wb[SHEET_BBG]["A1"].value == "Metal"
    assert wb[SHEET_BBG]["B2"].value == 9001.5
    # 原始數據應含未過濾的長天期列（含表頭 > 27 列）
    raw_rows = wb[SHEET_RAW].max_row
    assert raw_rows >= 40
    raw_a1 = str(wb[SHEET_RAW]["A1"].value or "")
    assert raw_a1.startswith("=")
    assert "20260819.xlsx" in raw_a1
    assert _winish(vba_dir) in _winish(raw_a1)
    print_ws = wb[SHEET_PRINT]
    assert "LME每日報價" in str(print_ws["A2"].value)
    assert "LME Daily Quotation" in str(print_ws["A3"].value)
    assert "第 &P 頁，共 &N 頁" in (print_ws.oddFooter.center.text or "")
    assert print_ws.page_setup.orientation == "landscape"
    if engine == "matplotlib":
        assert print_ws.page_setup.fitToWidth == 1
        assert len(wb[SHEET_CHART]._images) == 6
    wb.close()


def test_plot_window_truncates_long_tenor(tmp_path: Path):
    step2 = tmp_path / "curve.xlsx"
    _make_step2(step2)
    df = load_forward_curve(step2)
    plot_df = slice_plot_window(df, forward_months=27)
    cash = plot_df[DATE_COL].min()
    cutoff = cash + pd.DateOffset(months=27)
    assert plot_df[DATE_COL].max() <= cutoff
    assert len(plot_df) < len(df.dropna(subset=[DATE_COL]))
    # 各品種 dropna 不可把空值填 0
    ni = plot_df.dropna(subset=["NI"])
    assert ni["NI"].isna().sum() == 0
    assert (ni["NI"] == 0).sum() == 0
    assert set(SIX_SERIES) <= set(df.columns)


def test_excel_absolute_external_formula_uses_folder_and_file(tmp_path: Path):
    source = tmp_path / "work" / "20260820" / "20260820.xlsx"
    formula = excel_absolute_external_formula(source, "Sheet1", "A1")
    assert formula.startswith("='")
    assert "[20260820.xlsx]Sheet1" in formula
    assert formula.endswith("!A1")
    assert _winish(source.parent) in _winish(formula)


def test_missing_step2_raises(tmp_path: Path):
    cfg_path = _write_min_config(tmp_path, "matplotlib")
    config = load_config(cfg_path)
    as_of = date(2026, 8, 19)
    with pytest.raises(ReportBuildError, match="步驟二產出檔不存在"):
        build_report(
            config,
            as_of=as_of,
            step2_path=config.step2_workbook_path(as_of),
            bbg_values=(),
            bbg_formats=(),
        )


def test_output_dir_does_not_move_vba_or_chart_source(tmp_path: Path):
    import yaml

    work = tmp_path / "work"
    export = tmp_path / "LME_Export"
    work.mkdir()
    payload = {
        "paths": {
            "working_dir": str(work),
            "ref_workbook_name": "ref.xlsm",
            "bbg_workbook_name": "bbg.xlsx",
            "output_prefix": "LME每日報價",
            "output_dir": str(export),
        },
        "vba": {"macro_name": "RunDailyLME", "use_param_injection": True},
        "bloomberg": {
            "copy_range": "B3:I10",
            "bbg_sheet_name": "Promt date",
            "refresh_wait_seconds": 1,
        },
        "chart": {"forward_months": 27, "engine": "matplotlib"},
        "logging": {"level": "INFO", "file": ""},
    }
    cfg = tmp_path / "config.yaml"
    cfg.write_text(yaml.safe_dump(payload, allow_unicode=True), encoding="utf-8")
    config = load_config(cfg)
    as_of = date(2026, 8, 20)
    vba_dir, run_dir = config.ensure_run_dirs(as_of)
    assert vba_dir == work.resolve() / "20260820"
    assert run_dir == export.resolve() / "20260820"
    step2 = vba_dir / "20260820.xlsx"
    _make_step2(step2)
    dest = build_report(
        config,
        as_of=as_of,
        step2_path=step2,
        bbg_values=_bbg_sample()[0],
        bbg_formats=_bbg_sample()[1],
    )
    assert dest.parent == run_dir
    assert dest.is_file()
    assert step2.is_file()
    assert not (run_dir / "20260820.xlsx").exists()

    wb = load_workbook(dest)
    formula = str(wb[SHEET_RAW]["A1"].value or "")
    chart_src = str(wb[SHEET_CHART]["A2"].value or "")
    wb.close()
    assert formula.startswith("=")
    assert _winish(vba_dir) in _winish(formula)
    assert "20260820.xlsx" in formula
    assert _winish(run_dir) not in _winish(formula).replace(_winish(vba_dir), "")
    assert _winish(vba_dir) in _winish(chart_src)
    assert _winish(run_dir) not in _winish(chart_src)
