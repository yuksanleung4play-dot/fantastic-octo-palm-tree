"""把對映磁碟機（P:）路徑改成 UNC，避免 Excel 進程看不到磁碟機字母。"""

from __future__ import annotations

import re

# 與 config.yaml working_dir 同一台 share：P: → \\192.168.89.167\Dealing
P_DRIVE_UNC_ROOT = r"\\192.168.89.167\Dealing"

_P_DRIVE_RE = re.compile(r"(?i)(?<![A-Za-z0-9_])P:([\\/])")


def rewrite_p_drive_to_unc(text: str, *, unc_root: str = P_DRIVE_UNC_ROOT) -> str:
    """``P:\\Dealing Department - New\\...`` → ``\\\\192.168.89.167\\Dealing\\Dealing Department - New\\...``。"""
    if not text or "P:" not in text.upper():
        return text
    root = str(unc_root).rstrip("\\/")

    def _repl(match: re.Match[str]) -> str:
        sep = match.group(1)
        if sep == "/":
            return root.replace("\\", "/") + "/"
        return root + "\\"

    return _P_DRIVE_RE.sub(_repl, text)
