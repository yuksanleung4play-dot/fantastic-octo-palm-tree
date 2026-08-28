"""P: 對映磁碟機 ↔ UNC，以及 LME SPAN ``lme.yyyymmdd.dat`` 路徑。

Excel QueryTables 的 ``TEXT;`` 連線對 UNC 支援很差，常在檔案其實存在時
回報 1004 找不到 ``.dat``。``P:\\...`` 才能穩定 Refresh。上一版在跑巨集前
把 P: 改成 UNC，正是修改後找不到 dat 的主因。
"""

from __future__ import annotations

import re
from collections.abc import Iterable
from pathlib import Path

# 與 config.yaml working_dir 同一台 share：P: → \\192.168.89.167\Dealing
P_DRIVE_UNC_ROOT = r"\\192.168.89.167\Dealing"
DEFAULT_SPAN_FOLDER_NAME = "LME SPAN"

_P_DRIVE_RE = re.compile(r"(?i)(?<![A-Za-z0-9_])P:([\\/])")
_PADDED_NUMBERED_FOLDER_RE = re.compile(r"\d+\.\s{2,}")


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


def rewrite_unc_to_p_drive(text: str, *, unc_root: str = P_DRIVE_UNC_ROOT) -> str:
    """上一版 UNC 改寫的反向：``\\\\192.168.89.167\\Dealing\\...`` → ``P:\\...``。

    QueryTables ``TEXT;`` 連線需要磁碟機代號；檔案仍在同一台 share 上。
    """
    if not text:
        return text
    root = str(unc_root).rstrip("\\/")
    variants = [
        root,
        root.replace("\\", "/"),
        "//" + root.lstrip("\\/").replace("\\", "/"),
    ]
    out = text
    for variant in variants:
        if not variant:
            continue
        pattern = re.compile(re.escape(variant), re.IGNORECASE)
        out = pattern.sub("P:", out)
    return out


def default_span_dat_dir(working_dir: Path) -> Path:
    """``working_dir`` 是 ``...\\2. 期貨\\LME --Form & Sheet`` 時，SPAN 在同一層 ``LME SPAN``。"""
    return Path(working_dir).parent / DEFAULT_SPAN_FOLDER_NAME


def span_dat_filename(yyyymmdd: str) -> str:
    return f"lme.{yyyymmdd}.dat"


def span_dat_candidates(span_dir: Path, stamps: Iterable[str]) -> list[Path]:
    seen: list[str] = []
    for stamp in stamps:
        text = str(stamp).strip()
        if text and text not in seen:
            seen.append(text)
    return [Path(span_dir) / span_dat_filename(stamp) for stamp in seen]


def working_dir_has_padded_numbered_folders(path: Path | str) -> bool:
    """``1.      交易部`` 這種「數字.」後多個空白，常是 YAML 對齊誤加的。"""
    return _PADDED_NUMBERED_FOLDER_RE.search(str(path)) is not None
