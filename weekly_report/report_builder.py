"""把發件箱郵件與額外輸入資料，組裝成一份週報。

核心邏輯：
    - 將郵件依「寄出日期」分組，作為當天的日常工作內容。
    - 對主旨做去前綴（Re:／Fwd:）與去重，避免同一串對話重複佔行。
    - 合併使用者手動補充的工作項目。
    - 帶入「其他事項」。
"""

from __future__ import annotations

import re
from collections import OrderedDict
from datetime import date, timedelta
from typing import Dict, Iterable, List

from .models import EmailMessage, ReportConfig, WeeklyReport, WorkDay, WorkItem

# 主旨常見的回覆／轉寄前綴（中英文）。
_PREFIX_RE = re.compile(
    r"^\s*(re|fw|fwd|回复|回復|答复|轉發|转发|轉寄|轉發)\s*[:：]\s*",
    re.IGNORECASE,
)


def _normalize_subject(subject: str) -> str:
    """移除回覆／轉寄前綴，方便把同一個主題的郵件視為同一件事。"""

    text = subject.strip()
    # 可能有多層前綴，例如 "Re: Fwd: 主旨"。
    while True:
        new_text = _PREFIX_RE.sub("", text)
        if new_text == text:
            break
        text = new_text
    return text.strip()


def _list_subjects(messages: Iterable[EmailMessage]) -> List[WorkItem]:
    """把每封郵件分別列成一條工作項目（不做合併）。"""

    items: List[WorkItem] = []
    for msg in messages:
        subject = _normalize_subject(msg.subject) or "(無主旨)"
        items.append(
            WorkItem(
                description=subject,
                source="email",
                recipients=list(msg.recipients),
            )
        )
    return items


def _dedupe_subjects(messages: Iterable[EmailMessage]) -> List[WorkItem]:
    """同一天裡，把同主題（去前綴後相同）的郵件合併成一條工作項目。"""

    grouped: "OrderedDict[str, List[str]]" = OrderedDict()
    for msg in messages:
        key = _normalize_subject(msg.subject).lower()
        if not key:
            key = "(無主旨)"
        grouped.setdefault(key, [])
        for r in msg.recipients:
            if r not in grouped[key]:
                grouped[key].append(r)

    # 保留第一次出現時的原始主旨（去前綴）做為顯示文字。
    display: "OrderedDict[str, str]" = OrderedDict()
    for msg in messages:
        key = _normalize_subject(msg.subject).lower() or "(無主旨)"
        if key not in display:
            display[key] = _normalize_subject(msg.subject) or "(無主旨)"

    items: List[WorkItem] = []
    for key, recipients in grouped.items():
        items.append(
            WorkItem(
                description=display.get(key, key),
                source="email",
                recipients=recipients,
            )
        )
    return items


def _date_range(start: date, end: date) -> List[date]:
    days = []
    cur = start
    while cur <= end:
        days.append(cur)
        cur += timedelta(days=1)
    return days


def build_report(
    config: ReportConfig,
    emails: Iterable[EmailMessage],
    include_empty_days: bool = False,
    merge_threads: bool = False,
) -> WeeklyReport:
    """組裝週報。

    每日的「日常工作內容」由兩部分組成：

    1. **郵件記錄**：把當天發件箱的每一封郵件「分別列出」成一條工作內容
       （預設不合併；若 ``merge_threads=True`` 則把同主題對話合併成一條）。
    2. **自行輸入**：透過 ``config.manual_items`` 提供的每日工作內容
       （見 :meth:`weekly_report.models.ReportConfig.add_daily_work`），
       這就是讓使用者自由補充每日工作的「接口」。

    參數
    ----
    config:
        基本資訊與額外輸入（姓名、部門、期間、其他事項、每日工作內容）。
    emails:
        發件箱郵件（會自動依期間過濾）。
    include_empty_days:
        是否在報告中保留沒有任何工作項目的日期（預設略過）。
    merge_threads:
        是否把同一天同主題（去 Re:/Fwd: 後相同）的郵件合併成一條
        （預設 False，即每封郵件分別列出）。
    """

    by_day: Dict[date, List[EmailMessage]] = {}
    for msg in emails:
        if config.period_start <= msg.sent_date <= config.period_end:
            by_day.setdefault(msg.sent_date, []).append(msg)

    work_days: List[WorkDay] = []
    for day in _date_range(config.period_start, config.period_end):
        work_day = WorkDay(day=day)

        # (1) 來自郵件的工作項目：預設每封分別列出。
        day_emails = sorted(by_day.get(day, []), key=lambda m: m.sent_at)
        email_items = (
            _dedupe_subjects(day_emails) if merge_threads else _list_subjects(day_emails)
        )
        for item in email_items:
            work_day.add(item)

        # (2) 來自使用者自行輸入的每日工作內容。
        for text in _manual_for_day(config, day):
            work_day.add(WorkItem(description=text, source="manual"))

        if work_day.items or include_empty_days:
            work_days.append(work_day)

    return WeeklyReport(config=config, work_days=work_days)


def _manual_for_day(config: ReportConfig, day: date) -> List[str]:
    """從 config.manual_items 取出某天的手動項目。

    支援 key 為 ``datetime.date`` 或字串（YYYY-MM-DD）。
    """

    results: List[str] = []
    for key, values in config.manual_items.items():
        key_date = key
        if isinstance(key, str):
            try:
                from datetime import datetime as _dt

                key_date = _dt.strptime(key.strip(), "%Y-%m-%d").date()
            except ValueError:
                continue
        if key_date == day:
            if isinstance(values, str):
                results.append(values)
            else:
                results.extend(str(v) for v in values)
    return results
