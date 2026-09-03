"""一鍵腳本：預設綁定同目錄 config.yaml，可選自動開啟產出檔。"""

from __future__ import annotations

from datetime import date
from pathlib import Path

import yaml

from generate_lme_daily import _split_open_flag, main as generate_main
from tests.test_report_builder import _make_step2


def test_split_open_flag_defaults_to_open():
    rest, open_report = _split_open_flag(["--as-of", "2026-08-19"])
    assert rest == ["--as-of", "2026-08-19"]
    assert open_report is True


def test_split_open_flag_no_open():
    rest, open_report = _split_open_flag(["--no-open", "--dry-run"])
    assert rest == ["--dry-run"]
    assert open_report is False


def test_generate_dry_run(tmp_path: Path, capsys):
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
    code = generate_main(["--config", str(cfg), "--dry-run", "--no-open", "--as-of", "2026-08-19"])
    assert code == 0
    assert "dry-run" in capsys.readouterr().out.lower()


def test_generate_skip_vba_bbg_writes_named_report(tmp_path: Path):
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
    code = generate_main(
        [
            "--config",
            str(cfg),
            "--as-of",
            "2026-08-19",
            "--skip-vba",
            "--skip-bbg",
            "--no-open",
        ]
    )
    assert code == 0
    dest = work / "20260819" / "LME每日報價20260819.xlsx"
    assert dest.is_file()
    assert dest.stat().st_size > 0
    assert (work / "20260819" / "20260819.xlsx").is_file()
    assert list(dest.parent.glob("*.pdf")) == []
