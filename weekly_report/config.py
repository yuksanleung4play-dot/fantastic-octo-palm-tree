"""載入與建立 :class:`~weekly_report.models.ReportConfig`。

設定檔可以是 YAML（需安裝 PyYAML）或 JSON。範例見專案根目錄
``config.example.yaml``。
"""

from __future__ import annotations

import json
import os
from datetime import date, datetime, timedelta
from typing import Any, Dict, Optional

from .models import ReportConfig


def parse_date(value: Any) -> date:
    if isinstance(value, date) and not isinstance(value, datetime):
        return value
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, str):
        return datetime.strptime(value.strip(), "%Y-%m-%d").date()
    raise ValueError(f"無法解析日期：{value!r}")


def week_range(reference: Optional[date] = None, offset_weeks: int = 0) -> tuple:
    """回傳某一週的週一與週日。

    offset_weeks=0 為本週，-1 為上週。
    """

    today = reference or date.today()
    monday = today - timedelta(days=today.weekday()) + timedelta(weeks=offset_weeks)
    sunday = monday + timedelta(days=6)
    return monday, sunday


def _load_raw(path: str) -> Dict[str, Any]:
    if not os.path.exists(path):
        raise FileNotFoundError(f"找不到設定檔：{path}")

    with open(path, "r", encoding="utf-8") as fh:
        content = fh.read()

    if path.lower().endswith((".yaml", ".yml")):
        try:
            import yaml  # type: ignore
        except ImportError as exc:  # pragma: no cover - 取決於環境
            raise RuntimeError(
                "讀取 YAML 設定需要 PyYAML，請執行：pip install pyyaml，"
                "或改用 JSON 設定檔。"
            ) from exc
        return yaml.safe_load(content) or {}
    return json.loads(content)


def config_from_dict(data: Dict[str, Any]) -> ReportConfig:
    """由 dict 建立 ReportConfig，並套用合理的預設值。"""

    name = str(data.get("name", "")).strip()
    department = str(data.get("department", "")).strip()

    # 期間：可直接給 start/end，或用 week_offset 取某一週。
    if data.get("period_start") and data.get("period_end"):
        start = parse_date(data["period_start"])
        end = parse_date(data["period_end"])
    else:
        start, end = week_range(offset_weeks=int(data.get("week_offset", 0)))

    other_matters = data.get("other_matters") or []
    if isinstance(other_matters, str):
        other_matters = [other_matters]
    other_matters = [str(m).strip() for m in other_matters if str(m).strip()]

    manual_items = data.get("manual_items") or {}
    if not isinstance(manual_items, dict):
        manual_items = {}

    return ReportConfig(
        name=name,
        department=department,
        period_start=start,
        period_end=end,
        other_matters=other_matters,
        manual_items=manual_items,
        title=str(data.get("title", "週報")).strip() or "週報",
    )


def load_config(path: str) -> ReportConfig:
    return config_from_dict(_load_raw(path))


def email_settings(path: str) -> Dict[str, Any]:
    """從設定檔取出 email 區塊（IMAP / 本地檔）。"""

    raw = _load_raw(path)
    return raw.get("email") or {}
