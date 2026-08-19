"""config 載入與路徑檢查。"""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pytest
import yaml

from lme_daily.config import load_config, relax_windows_yaml_quotes, validate_required_files
from lme_daily.exceptions import ConfigError


def _write_config(tmp_path: Path, **overrides) -> Path:
    payload = {
        "paths": {
            "working_dir": str(tmp_path / "work"),
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
        "holidays": {"dates": ["2026-12-25"]},
        "logging": {"level": "INFO", "file": ""},
    }
    for key, value in overrides.items():
        if isinstance(value, dict) and key in payload:
            payload[key].update(value)
        else:
            payload[key] = value
    path = tmp_path / "config.yaml"
    path.write_text(yaml.safe_dump(payload, allow_unicode=True), encoding="utf-8")
    return path


def test_load_config_combines_paths(tmp_path: Path):
    work = tmp_path / "work"
    work.mkdir()
    cfg_path = _write_config(tmp_path)
    config = load_config(cfg_path)
    assert config.paths.working_dir == work.resolve()
    assert config.paths.ref_workbook == (work / "ref.xlsm").resolve()
    assert config.paths.bbg_workbook == (work / "bbg.xlsx").resolve()
    assert config.vba.macro_name == "RunDailyLME"
    assert config.chart.forward_months == 27
    assert date(2026, 12, 25) in config.holidays


def test_validate_required_files_fails_when_missing(tmp_path: Path):
    work = tmp_path / "work"
    work.mkdir()
    cfg_path = _write_config(tmp_path)
    config = load_config(cfg_path)
    with pytest.raises(ConfigError, match="參考工作簿不存在"):
        validate_required_files(config, require_workbooks=True)


def test_validate_passes_when_workbooks_exist(tmp_path: Path):
    work = tmp_path / "work"
    work.mkdir()
    (work / "ref.xlsm").write_bytes(b"fake")
    (work / "bbg.xlsx").write_bytes(b"fake")
    cfg_path = _write_config(tmp_path)
    config = load_config(cfg_path)
    validate_required_files(config, require_workbooks=True)


def test_invalid_chart_engine(tmp_path: Path):
    (tmp_path / "work").mkdir()
    cfg_path = _write_config(tmp_path, chart={"forward_months": 27, "engine": "png"})
    with pytest.raises(ConfigError, match="chart.engine"):
        load_config(cfg_path)


def test_missing_config_file(tmp_path: Path):
    with pytest.raises(ConfigError, match="指定的設定檔不存在"):
        load_config(tmp_path / "nope.yaml")


def test_relax_turns_double_quoted_unc_into_single_quotes():
    src = 'working_dir: "\\\\192.168.89.167\\Dealing\\Dept"\n'
    with pytest.raises(yaml.YAMLError, match="unknown escape"):
        yaml.safe_load(src)
    fixed = relax_windows_yaml_quotes(src)
    data = yaml.safe_load(fixed)
    assert data["working_dir"] == r"\\192.168.89.167\Dealing\Dept"


def test_load_config_accepts_unc_written_like_windows_users(tmp_path: Path):
    """Users write "\\\\server\\Dealing\\..." in double quotes; vanilla YAML crashes."""
    body = "\n".join(
        [
            "paths:",
            '  working_dir: "\\\\192.168.89.167\\Dealing\\Dept"',
            "  ref_workbook_name: ref.xlsm",
            "  bbg_workbook_name: bbg.xlsx",
            "  output_prefix: LME每日報價",
            "vba:",
            "  macro_name: RunDailyLME",
            "  use_param_injection: true",
            "bloomberg:",
            "  copy_range: B3:I10",
            "  bbg_sheet_name: Promt date",
            "  refresh_wait_seconds: 1",
            "chart:",
            "  forward_months: 27",
            "  engine: matplotlib",
            "logging:",
            "  level: INFO",
            '  file: ""',
        ]
    )
    cfg = tmp_path / "config.yaml"
    cfg.write_text(body + "\n", encoding="utf-8")
    config = load_config(cfg)
    assert "192.168.89.167" in str(config.paths.working_dir)
    assert "Dealing" in str(config.paths.working_dir)
