"""Windows 主控台編碼與相依套件檢查（一鍵啟動用）。"""

from __future__ import annotations

import importlib
import os
import subprocess
import sys

CORE_MODULES = (
    ("yaml", "pyyaml"),
    ("dateutil", "python-dateutil"),
    ("pandas", "pandas"),
    ("openpyxl", "openpyxl"),
    ("xlsxwriter", "xlsxwriter"),
    ("matplotlib", "matplotlib"),
)

WIN_MODULES = (
    ("win32com.client", "pywin32"),
)


def configure_windows_console() -> None:
    """避免 cmd.exe CP950 把 Python 中文印成亂碼。"""
    if sys.platform != "win32":
        return
    os.environ.setdefault("PYTHONUTF8", "1")
    os.environ.setdefault("PYTHONIOENCODING", "utf-8")
    try:
        import ctypes

        ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        ctypes.windll.kernel32.SetConsoleCP(65001)
    except Exception:
        pass
    for name in ("stdout", "stderr"):
        stream = getattr(sys, name, None)
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            try:
                reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


def _missing_packages(*, need_win32: bool) -> list[str]:
    missing: list[str] = []
    for module_name, pip_name in CORE_MODULES:
        try:
            importlib.import_module(module_name)
        except ImportError:
            missing.append(pip_name)
    if need_win32 and sys.platform == "win32":
        for module_name, pip_name in WIN_MODULES:
            try:
                importlib.import_module(module_name)
            except ImportError:
                missing.append(pip_name)
    return missing


def ensure_dependencies(argv: list[str] | None = None, *, auto_install: bool = True) -> None:
    """缺少套件時自動 ``pip install``；失敗則印出可複製的 ASCII 指令。"""
    argv = list(argv or [])
    dry = "--dry-run" in argv
    skip_vba = "--skip-vba" in argv
    skip_bbg = "--skip-bbg" in argv
    need_win32 = (not dry) and (not skip_vba or not skip_bbg)
    missing = _missing_packages(need_win32=need_win32)
    if not missing:
        return

    unique = list(dict.fromkeys(missing))
    print("Missing packages: " + ", ".join(unique))
    cmd = [sys.executable, "-m", "pip", "install", *unique]
    if not auto_install:
        print("Run: " + " ".join(cmd))
        raise SystemExit(2)

    print("Installing: " + " ".join(cmd))
    try:
        subprocess.check_call(cmd)
    except subprocess.CalledProcessError as exc:
        print("pip install failed.")
        print("Run this in Command Prompt:")
        print("  " + " ".join(cmd))
        raise SystemExit(exc.returncode) from exc

    still = _missing_packages(need_win32=need_win32)
    if still:
        print("Still missing after pip install: " + ", ".join(still))
        if "pywin32" in still:
            print("Try:  python -m pip install pywin32")
            print("Then: python -m pywin32_postinstall -install")
        raise SystemExit(2)
    print("Packages installed OK.")
