"""週報生成器的單元測試。"""

import os
import sys
import unittest
from datetime import date, datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from weekly_report.config import config_from_dict, parse_date, week_range
from weekly_report.email_source import load_from_file
from weekly_report.models import EmailMessage, ReportConfig
from weekly_report.renderers import render
from weekly_report.report_builder import _normalize_subject, build_report

SAMPLE_MBOX = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "samples",
    "sent.mbox",
)


def make_email(subject, day, hour=9, recipients=None):
    return EmailMessage(
        subject=subject,
        sent_at=datetime(2026, 6, day, hour, 0),
        recipients=recipients or ["a@example.com"],
        sender="me@example.com",
    )


class NormalizeSubjectTests(unittest.TestCase):
    def test_strip_reply_prefix(self):
        self.assertEqual(_normalize_subject("Re: 主旨"), "主旨")
        self.assertEqual(_normalize_subject("FW: 主旨"), "主旨")
        self.assertEqual(_normalize_subject("回复：主旨"), "主旨")

    def test_strip_nested_prefix(self):
        self.assertEqual(_normalize_subject("Re: Fwd: 主旨"), "主旨")

    def test_no_prefix(self):
        self.assertEqual(_normalize_subject("正常主旨"), "正常主旨")


class BuildReportTests(unittest.TestCase):
    def setUp(self):
        self.config = ReportConfig(
            name="測試員",
            department="測試部",
            period_start=date(2026, 6, 1),
            period_end=date(2026, 6, 7),
            other_matters=["事項一"],
        )

    def test_groups_by_day(self):
        emails = [
            make_email("任務A", 1),
            make_email("任務B", 1, hour=14),
            make_email("任務C", 3),
        ]
        report = build_report(self.config, emails)
        self.assertEqual(len(report.work_days), 2)
        self.assertEqual(report.work_days[0].day, date(2026, 6, 1))
        self.assertEqual(len(report.work_days[0].items), 2)

    def test_merges_reply_threads_same_day(self):
        emails = [
            make_email("專案啟動", 1),
            make_email("Re: 專案啟動", 1, hour=15),
        ]
        report = build_report(self.config, emails)
        self.assertEqual(len(report.work_days[0].items), 1)
        self.assertEqual(report.work_days[0].items[0].description, "專案啟動")

    def test_filters_outside_period(self):
        emails = [make_email("超出範圍", 30)]  # 6/30 不在 6/1~6/7
        report = build_report(self.config, emails)
        self.assertEqual(report.total_items, 0)

    def test_manual_items_merged(self):
        self.config.manual_items = {"2026-06-02": ["手動工作"]}
        report = build_report(self.config, [make_email("郵件工作", 2)])
        day = report.day_for(date(2026, 6, 2))
        descriptions = [i.description for i in day.items]
        self.assertIn("手動工作", descriptions)
        self.assertIn("郵件工作", descriptions)

    def test_include_empty_days(self):
        report = build_report(self.config, [], include_empty_days=True)
        self.assertEqual(len(report.work_days), 7)


class RendererTests(unittest.TestCase):
    def setUp(self):
        config = ReportConfig(
            name="測試員",
            department="測試部",
            period_start=date(2026, 6, 1),
            period_end=date(2026, 6, 7),
            other_matters=["其他事項一"],
        )
        self.report = build_report(config, [make_email("某工作", 1)])

    def test_text_contains_sections(self):
        out = render(self.report, "text")
        self.assertIn("測試員", out)
        self.assertIn("日常工作內容", out)
        self.assertIn("其他事項", out)

    def test_markdown(self):
        out = render(self.report, "markdown")
        self.assertIn("# 週報", out)
        self.assertIn("某工作", out)

    def test_html_escapes(self):
        config = ReportConfig(
            name="<b>x</b>",
            department="d",
            period_start=date(2026, 6, 1),
            period_end=date(2026, 6, 7),
        )
        report = build_report(config, [])
        out = render(report, "html")
        self.assertIn("&lt;b&gt;", out)
        self.assertNotIn("<b>x</b>", out)

    def test_unknown_format(self):
        with self.assertRaises(ValueError):
            render(self.report, "pdf")


class ConfigTests(unittest.TestCase):
    def test_parse_date(self):
        self.assertEqual(parse_date("2026-06-01"), date(2026, 6, 1))

    def test_week_range(self):
        start, end = week_range(reference=date(2026, 6, 3))  # 週三
        self.assertEqual(start, date(2026, 6, 1))
        self.assertEqual(end, date(2026, 6, 7))

    def test_config_from_dict_defaults(self):
        cfg = config_from_dict({"name": "甲", "department": "乙"})
        self.assertEqual(cfg.name, "甲")
        self.assertEqual(cfg.title, "週報")
        self.assertEqual((cfg.period_end - cfg.period_start).days, 6)

    def test_other_matters_string(self):
        cfg = config_from_dict({"name": "甲", "other_matters": "單一事項"})
        self.assertEqual(cfg.other_matters, ["單一事項"])


class EmailSourceTests(unittest.TestCase):
    def test_load_sample_mbox(self):
        emails = load_from_file(SAMPLE_MBOX, date(2026, 6, 1), date(2026, 6, 7))
        self.assertEqual(len(emails), 5)
        self.assertEqual(emails[0].sent_date, date(2026, 6, 1))
        # 中文主旨正確解碼。
        self.assertTrue(any("報表" in e.subject for e in emails))

    def test_load_missing_file(self):
        with self.assertRaises(FileNotFoundError):
            load_from_file("/nonexistent/path.mbox", date(2026, 6, 1), date(2026, 6, 7))


if __name__ == "__main__":
    unittest.main()
