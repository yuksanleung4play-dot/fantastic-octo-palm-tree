"""週報生成器 (Weekly Report Generator).

根據郵箱發件箱與額外輸入的資料，整理成一份結構化的週報。

週報格式：
    1) 姓名、日期、部門
    2) 日常工作內容
    3) 其他事項
"""

from .models import EmailMessage, ReportConfig, WeeklyReport, WorkDay, WorkItem

__all__ = [
    "EmailMessage",
    "ReportConfig",
    "WeeklyReport",
    "WorkDay",
    "WorkItem",
    "__version__",
]

__version__ = "1.0.0"
