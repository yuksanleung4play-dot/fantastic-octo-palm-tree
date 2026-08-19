"""dry-run 串接：不算 Excel COM。"""

from __future__ import annotations

from datetime import date
from pathlib import Path

import yaml

from lme_daily.config import load_config
from lme_daily.main import run
from lme_daily.report_builder import SHEET_BBG, SHEET_CHART, SHEET_RAW
from tests.test_report_builder import _make_step2


def test_dry_run_returns_none(tmp_path: Path):
    work = tmp_path / "work"
    work.mkdir()
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
        "chart": {"forward_months": 27, "engine": "matplotlib"},
        "logging": {"level": "INFO", "file": ""},
    }
    cfg = tmp_path / "config.yaml"
    cfg.write_text(yaml.safe_dump(payload, allow_unicode=True), encoding="utf-8")
    config = load_config(cfg)
    assert run(config, as_of=date(2026, 8, 19), dry_run=True, skip_vba=False, skip_bbg=False) is None


def test_skip_vba_and_bbg_builds_report(tmp_path: Path):
    work = tmp_path / "work"
    work.mkdir()
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
        "chart": {"forward_months": 27, "engine": "matplotlib"},
        "logging": {"level": "INFO", "file": ""},
    }
    cfg = tmp_path / "config.yaml"
    cfg.write_text(yaml.safe_dump(payload, allow_unicode=True), encoding="utf-8")
    _make_step2(work / "20260819.xlsx")
    config = load_config(cfg)
    dest = run(
        config,
        as_of=date(2026, 8, 19),
        dry_run=False,
        skip_vba=True,
        skip_bbg=True,
    )
    assert dest is not None
    assert dest.is_file()
    from openpyxl import load_workbook

    wb = load_workbook(dest)
    assert {SHEET_BBG, SHEET_CHART, SHEET_RAW} <= set(wb.sheetnames)
    wb.close()
