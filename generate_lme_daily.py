#!/usr/bin/env python3
"""One-click generator for LME daily report xlsx.

Windows: double-click RUN_LME.bat  (ASCII, no UTF-8 BOM)

    python generate_lme_daily.py
    python generate_lme_daily.py --no-open
    python generate_lme_daily.py --as-of 2026-08-19
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def _split_open_flag(argv: list[str]) -> tuple[list[str], bool]:
    open_report = True
    kept: list[str] = []
    for arg in argv:
        if arg in {"--no-open", "--no-open-report"}:
            open_report = False
        elif arg in {"--open", "--open-report"}:
            open_report = True
        else:
            kept.append(arg)
    return kept, open_report


def open_report_file(path: Path) -> None:
    resolved = path.resolve()
    if not resolved.is_file():
        print(f"ERROR: report not found: {resolved}", file=sys.stderr)
        return
    try:
        if sys.platform == "win32":
            os.startfile(str(resolved))  # type: ignore[attr-defined]
        else:
            import subprocess

            opener = "open" if sys.platform == "darwin" else "xdg-open"
            subprocess.Popen([opener, str(resolved)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        print(f"Opened: {resolved}")
    except OSError as exc:
        print(f"Created file but could not open it: {exc}", file=sys.stderr)
        print(f"Open manually: {resolved}", file=sys.stderr)


def main(argv: list[str] | None = None) -> int:
    from lme_daily.bootstrap import configure_windows_console, ensure_dependencies
    from lme_daily.main import run_cli

    configure_windows_console()

    raw = list(sys.argv[1:] if argv is None else argv)
    rest, open_report = _split_open_flag(raw)
    if "--config" not in rest and "-h" not in rest and "--help" not in rest:
        default_config = ROOT / "config.yaml"
        rest = ["--config", str(default_config), *rest]
        if not default_config.is_file() and "--help" not in rest:
            print("ERROR: config.yaml not found next to generate_lme_daily.py")
            print("Expected: " + str(default_config))
            return 1

    auto_install = os.environ.get("LME_SKIP_PIP") != "1"
    if "-h" not in rest and "--help" not in rest:
        ensure_dependencies(rest, auto_install=auto_install)

    print("=" * 60)
    print("  Generate LME daily report  LME每日報價yyyymmdd.xlsx")
    print("=" * 60)
    print("Script folder: " + str(ROOT))

    prev_cwd = Path.cwd()
    os.chdir(ROOT)
    try:
        code, output = run_cli(rest)
    finally:
        os.chdir(prev_cwd)
    if code != 0:
        print()
        print("FAILED. Read the error above.")
        print("Checklist:")
        print("  1) Edit config.yaml paths.working_dir")
        print("  2) Excel installed, Bloomberg Terminal logged in")
        print("  3) Source workbooks exist in working_dir")
        return code

    if output is None:
        print()
        print("dry-run finished (no xlsx written).")
        return 0

    print()
    print("OUTPUT=")
    print(str(output.resolve()))
    if open_report:
        open_report_file(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
