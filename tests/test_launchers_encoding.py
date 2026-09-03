"""Launchers must be ASCII without UTF-8 BOM (CP950 cmd misreads BOM as 嘿濃)."""

from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def test_batch_files_have_no_utf8_bom_and_are_ascii():
    files = [
        ROOT / "RUN_LME.bat",
        ROOT / "Generate_LME_Daily.bat",
        ROOT / "install_deps.bat",
        ROOT / "一鍵生成LME每日報價.bat",
        ROOT / "RUN_LME.vbs",
        ROOT / "HOW_TO_RUN.txt",
    ]
    for path in files:
        data = path.read_bytes()
        assert not data.startswith(b"\xef\xbb\xbf"), f"UTF-8 BOM in {path.name}"
        data.decode("ascii")  # raises if non-ASCII
        if path.suffix.lower() == ".bat":
            assert data.lstrip(b"\r\n").startswith(b"@echo off"), path.name


def test_run_lme_bat_uses_relative_python_after_pushd():
    text = (ROOT / "RUN_LME.bat").read_text(encoding="ascii")
    assert "pushd \"%~dp0\"" in text
    assert "chcp 65001" not in text
    assert "generate_lme_daily.py" in text
