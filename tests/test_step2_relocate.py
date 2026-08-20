"""VBA 中繼檔固定留在 vba_dir（不需 Excel COM）。"""

from __future__ import annotations

from datetime import date
from pathlib import Path

import yaml

from lme_daily.config import load_config
from lme_daily.vba_runner import relocate_step2_workbook, resolve_existing_step2, step2_search_paths


def test_relocate_moves_legacy_file_into_vba_dir(tmp_path: Path):
    src = tmp_path / "20260819.xlsx"
    src.write_bytes(b"curve")
    dest = tmp_path / "20260819" / "20260819.xlsx"
    ready = relocate_step2_workbook(src, dest)
    assert ready == dest
    assert dest.is_file()
    assert dest.read_bytes() == b"curve"
    assert not src.exists()


def test_relocate_same_path_is_noop(tmp_path: Path):
    dest = tmp_path / "20260819" / "20260819.xlsx"
    dest.parent.mkdir()
    dest.write_bytes(b"keep")
    assert relocate_step2_workbook(dest, dest) == dest
    assert dest.read_bytes() == b"keep"


def test_resolve_existing_prefers_vba_dir(tmp_path: Path):
    work = tmp_path / "work"
    work.mkdir()
    vba_dir = work / "20260819"
    vba_dir.mkdir()
    (vba_dir / "20260819.xlsx").write_bytes(b"vba")
    (work / "20260819.xlsx").write_bytes(b"legacy")
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
    as_of = date(2026, 8, 19)
    paths = step2_search_paths(config, as_of)
    assert paths[0] == config.step2_workbook_path(as_of)
    found = resolve_existing_step2(config, as_of)
    assert found == vba_dir / "20260819.xlsx"
