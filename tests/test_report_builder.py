"""報告產生（matplotlib / xlsxwriter），不需 Excel COM。"""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import load_workbook

from lme_daily.config import load_config
from lme_daily.report_builder import (
    DATE_COL,
    SHEET_BBG,
    SHEET_CHART,
    SHEET_RAW,
    SIX_SERIES,
    build_report,
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
    step2 = config.paths.working_dir / "20260819.xlsx"
    _make_step2(step2)
    values, formats = _bbg_sample()

    dest = build_report(
        config,
        as_of=date(2026, 8, 19),
        step2_path=step2,
        bbg_values=values,
        bbg_formats=formats,
    )
    assert dest.name == "LME每日報價20260819.xlsx"
    assert dest.is_file()

    wb = load_workbook(dest)
    assert SHEET_BBG in wb.sheetnames
    assert SHEET_CHART in wb.sheetnames
    assert SHEET_RAW in wb.sheetnames
    assert wb[SHEET_BBG]["A1"].value == "Metal"
    assert wb[SHEET_BBG]["B2"].value == 9001.5
    # 原始數據應含未過濾的長天期列（含表頭 > 27 列）
    raw_rows = wb[SHEET_RAW].max_row
    assert raw_rows >= 40
    if engine == "matplotlib":
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
