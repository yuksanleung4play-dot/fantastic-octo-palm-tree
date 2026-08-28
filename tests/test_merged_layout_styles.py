"""合併版面定案樣式：A3 主標題、C8 表頭代碼、A271 免責聲明。"""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import load_workbook

from lme_daily.config import load_config
from lme_daily.report_builder import (
    COLOR_ACCENT,
    COLOR_DATE_BG,
    COLOR_DATE_TEXT,
    COLOR_DISCLAIMER_TEXT,
    COLOR_NAVY,
    COLOR_TEXT_CREAM,
    COLOR_TEXT_WHITE,
    COLUMN_WIDTHS,
    DATE_COL,
    FONT_NAME,
    PRINT_DISCLAIMER_FULL,
    SHEET_BBG,
    SHEET_PRINT,
    build_report,
    merged_layout_disclaimer_row,
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


def _hex_color(color) -> str | None:
    if color is None:
        return None
    rgb = getattr(color, "rgb", None)
    if rgb is None or rgb in ("00000000", "None", "0"):
        return None
    text = str(rgb)
    if text in {"00000000", "None", "0"}:
        return None
    if len(text) >= 6:
        return text[-6:].upper()
    return text.upper()


def fill_hex(cell) -> str | None:
    fill = cell.fill
    if fill is None or fill.patternType in (None, "none"):
        return None
    return _hex_color(fill.fgColor)


def font_hex(cell) -> str | None:
    font = cell.font
    if font is None:
        return None
    return _hex_color(font.color)


def _bbg_tenor_sample():
    """對應 B3:I10 的 8×8 表頭代碼列（C:H = CA/AH/PB/NI/SN/ZS）。"""
    values = (
        ("", "", "CA", "AH", "PB", "NI", "SN", "ZS"),
        (
            "Tenors",
            "Prompt date",
            "Copper",
            "Primary Aluminium",
            "Lead",
            "Nickel",
            "Tin",
            "Zinc",
        ),
        ("Cash", date(2026, 9, 1), 14461.57, 3218.10, 1873.41, 16712.76, 54479.00, 4091.95),
        ("3MO", date(2026, 11, 27), 14253.00, 3225.00, 1910.50, 16887.00, 54827.00, 3893.00),
        ("15Mo", date(2027, 11, 17), 14171.68, 3146.00, 2051.03, 17619.50, 55040.00, 3722.13),
        ("27Mo", date(2028, 11, 15), 14138.43, 3072.00, 2144.28, 18241.50, "N/A", 3470.88),
        ("63Mo", date(2031, 11, 19), 14146.43, 2992.00, 2224.78, 20245.50, "N/A", 3154.88),
        ("123Mo", date(2036, 11, 19), 14148.43, 3094.00, "N/A", "N/A", "N/A", "N/A"),
    )
    formats = (("General",) * 8,) * 8
    return values, formats


def _make_step2_rows(path: Path, n_data: int) -> None:
    cash = datetime(2026, 8, 28)
    rows = []
    for i in range(n_data):
        d = cash + pd.Timedelta(days=i)
        rows.append(
            {
                DATE_COL: d,
                "CA": 9000 + i,
                "AH": 2500 + i,
                "ZS": 2800 + i,
                "NI": 16000 + i,
                "SN": 30000 + i,
                "PB": 2000 + i,
            }
        )
    pd.DataFrame(rows).to_excel(path, index=False)


def test_merged_layout_disclaimer_row_matches_fu_ben_a271():
    assert merged_layout_disclaimer_row(bbg_rows=8, raw_rows=210) == 271


@pytest.mark.parametrize("engine", ["matplotlib", "xlsxwriter"])
def test_merged_layout_frozen_styles_a3_c8_a271(tmp_path: Path, engine: str):
    cfg_path = _write_min_config(tmp_path, engine)
    config = load_config(cfg_path)
    as_of = date(2026, 8, 27)
    vba_dir = config.vba_dir(as_of)
    vba_dir.mkdir(parents=True, exist_ok=True)
    step2 = vba_dir / "20260827.xlsx"
    _make_step2_rows(step2, 209)
    values, formats = _bbg_tenor_sample()
    assert merged_layout_disclaimer_row(bbg_rows=len(values), raw_rows=210) == 271

    dest = build_report(
        config,
        as_of=as_of,
        step2_path=step2,
        bbg_values=values,
        bbg_formats=formats,
    )
    wb = load_workbook(dest)
    ws = wb[SHEET_PRINT]

    assert ws["A3"].value == "LME每日報價"
    assert fill_hex(ws["A3"]) == COLOR_ACCENT
    assert font_hex(ws["A3"]) == COLOR_TEXT_CREAM
    assert ws["A3"].font.bold is True
    assert float(ws["A3"].font.size) == 20
    assert ws["A3"].font.name == FONT_NAME
    assert float(ws.row_dimensions[3].height) == 34

    assert ws["C8"].value == "CA"
    assert fill_hex(ws["C8"]) == COLOR_ACCENT
    assert font_hex(ws["C8"]) == COLOR_TEXT_WHITE
    assert ws["C8"].font.bold is True
    assert ws["C8"].alignment.horizontal == "center"
    assert ws["C8"].alignment.vertical == "center"
    assert ws["H8"].value == "ZS"
    assert fill_hex(ws["H8"]) == COLOR_ACCENT
    assert font_hex(ws["H8"]) == COLOR_TEXT_WHITE

    assert fill_hex(ws["A5"]) == COLOR_DATE_BG
    assert font_hex(ws["A5"]) == COLOR_DATE_TEXT

    disc = ws["A271"]
    assert disc.value == PRINT_DISCLAIMER_FULL
    assert fill_hex(disc) is None
    assert font_hex(disc) == COLOR_DISCLAIMER_TEXT
    assert float(disc.font.size) == 8
    assert disc.alignment.wrap_text is True
    assert disc.alignment.horizontal == "left"
    assert disc.alignment.vertical == "top"
    height = ws.row_dimensions[271].height
    assert height is not None and float(height) >= 78

    def _width(letter: str) -> float:
        value = ws.column_dimensions[letter].width
        assert value is not None
        return float(value)

    assert _width("A") == pytest.approx(COLUMN_WIDTHS["A"], abs=1)
    assert _width("B") == pytest.approx(COLUMN_WIDTHS["B"], abs=1)
    assert _width("D") == pytest.approx(COLUMN_WIDTHS["D"], abs=1)
    assert _width("E") == pytest.approx(COLUMN_WIDTHS["E"], abs=1)

    bbg = wb[SHEET_BBG]
    assert fill_hex(bbg["A1"]) == COLOR_NAVY
    assert font_hex(bbg["A1"]) == COLOR_TEXT_WHITE
    wb.close()
