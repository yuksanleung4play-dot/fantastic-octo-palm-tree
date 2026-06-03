"""週報所使用的資料模型。

這裡定義純粹的資料結構（dataclass），不含任何 I/O 邏輯，
方便測試與在不同來源（IMAP、本地郵件檔、手動輸入）之間共用。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from typing import List, Optional


@dataclass
class EmailMessage:
    """一封已寄出的郵件（發件箱）。"""

    subject: str
    sent_at: datetime
    recipients: List[str] = field(default_factory=list)
    sender: str = ""
    snippet: str = ""

    @property
    def sent_date(self) -> date:
        return self.sent_at.date()

    @property
    def recipients_text(self) -> str:
        return ", ".join(r for r in self.recipients if r)


@dataclass
class WorkItem:
    """一條工作內容。

    可以由郵件自動推導，亦可由使用者手動補充。
    """

    description: str
    # 來源：'email' 代表由發件箱推導，'manual' 代表手動補充。
    source: str = "manual"
    recipients: List[str] = field(default_factory=list)

    @property
    def recipients_text(self) -> str:
        return ", ".join(r for r in self.recipients if r)


@dataclass
class WorkDay:
    """某一天的工作彙整。"""

    day: date
    items: List[WorkItem] = field(default_factory=list)

    WEEKDAY_ZH = ["週一", "週二", "週三", "週四", "週五", "週六", "週日"]

    @property
    def weekday_zh(self) -> str:
        return self.WEEKDAY_ZH[self.day.weekday()]

    def add(self, item: WorkItem) -> None:
        self.items.append(item)


@dataclass
class ReportConfig:
    """產生週報所需的基本資訊與額外輸入。"""

    name: str
    department: str
    # 週報涵蓋的起訖日期。
    period_start: date
    period_end: date
    # 「其他事項」：例如下週計劃、需要協調的事、請假等。
    other_matters: List[str] = field(default_factory=list)
    # 額外的手動工作項目，key 為日期，value 為內容字串列表。
    manual_items: dict = field(default_factory=dict)
    title: str = "週報"

    @property
    def period_text(self) -> str:
        return f"{self.period_start:%Y-%m-%d} ~ {self.period_end:%Y-%m-%d}"

    def add_daily_work(self, day, text: str) -> None:
        """每日日常工作內容的輸入接口。

        讓使用者（或其他程式）自行為某一天補充工作內容。可重複呼叫，
        同一天的多條內容會累加。

        參數
        ----
        day:
            ``datetime.date`` 或 ``YYYY-MM-DD`` 字串。
        text:
            該天的一條工作內容。
        """

        if isinstance(day, datetime):
            day = day.date()
        if isinstance(day, date):
            key = day.strftime("%Y-%m-%d")
        else:
            key = str(day).strip()

        if not text or not str(text).strip():
            return

        bucket = self.manual_items.setdefault(key, [])
        if isinstance(bucket, str):
            bucket = [bucket]
            self.manual_items[key] = bucket
        bucket.append(str(text).strip())


@dataclass
class WeeklyReport:
    """組裝完成、可供輸出的週報。"""

    config: ReportConfig
    work_days: List[WorkDay] = field(default_factory=list)
    generated_at: datetime = field(default_factory=datetime.now)

    @property
    def total_items(self) -> int:
        return sum(len(d.items) for d in self.work_days)

    def day_for(self, day: date) -> Optional[WorkDay]:
        for wd in self.work_days:
            if wd.day == day:
                return wd
        return None
