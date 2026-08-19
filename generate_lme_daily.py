#!/usr/bin/env python3
"""一鍵產生 ``LME每日報價{yyyymmdd}.xlsx``。

建議用同目錄的 ``一鍵生成LME每日報價.bat`` 雙擊執行（Windows）。
也可在命令列：

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
    """用系統預設程式（通常是 Excel）開啟產出檔。"""
    resolved = path.resolve()
    if not resolved.is_file():
        print(f"找不到報告檔，無法開啟：{resolved}", file=sys.stderr)
        return
    try:
        if sys.platform == "win32":
            os.startfile(str(resolved))  # type: ignore[attr-defined]
        else:
            import subprocess

            opener = "open" if sys.platform == "darwin" else "xdg-open"
            subprocess.Popen([opener, str(resolved)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        print(f"已開啟：{resolved}")
    except OSError as exc:
        print(f"已產生檔案，但自動開啟失敗：{exc}\n請自行開啟：{resolved}", file=sys.stderr)


def main(argv: list[str] | None = None) -> int:
    from lme_daily.main import run_cli

    raw = list(sys.argv[1:] if argv is None else argv)
    rest, open_report = _split_open_flag(raw)
    if "--config" not in rest and "-h" not in rest and "--help" not in rest:
        default_config = ROOT / "config.yaml"
        rest = ["--config", str(default_config), *rest]

    print("=" * 60)
    print("  一鍵生成 LME每日報價yyyymmdd.xlsx")
    print("=" * 60)

    prev_cwd = Path.cwd()
    os.chdir(ROOT)
    try:
        code, output = run_cli(rest)
    finally:
        os.chdir(prev_cwd)
    if code != 0:
        print()
        print("產生失敗，視窗可保留對照上方錯誤。")
        return code

    if output is None:
        print()
        print("dry-run 完成（未寫出 Excel）。")
        return 0

    print()
    print("已產生最終報告：")
    print(f"  {output.resolve()}")
    if open_report:
        open_report_file(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
