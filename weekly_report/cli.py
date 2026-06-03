"""命令列介面。

範例：

    # 用本地匯出的 .mbox 產生本週週報（Markdown）
    python -m weekly_report --mbox sent.mbox --name 王小明 --department 研發部 -f markdown

    # 用設定檔（含 IMAP 帳密與其他事項）
    python -m weekly_report --config config.yaml -o report.html -f html

    # 不接郵箱，純手動 / 互動輸入
    python -m weekly_report --name 王小明 --department 研發部 --interactive
"""

from __future__ import annotations

import argparse
import getpass
import os
import sys
from datetime import date
from typing import List, Optional

from . import __version__
from .config import config_from_dict, email_settings, load_config, parse_date, week_range
from .email_source import load_from_file, load_from_imap
from .models import EmailMessage, ReportConfig
from .renderers import RENDERERS, render
from .report_builder import build_report


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="weekly_report",
        description="根據郵箱發件箱與額外輸入，整理成一份週報。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("--config", help="設定檔路徑（.yaml 或 .json）")

    # 基本資訊（可覆蓋設定檔）。
    p.add_argument("--name", help="姓名")
    p.add_argument("--department", help="部門")
    p.add_argument("--title", help="週報標題（預設：週報）")

    # 期間。
    p.add_argument("--start", help="起始日期 YYYY-MM-DD")
    p.add_argument("--end", help="結束日期 YYYY-MM-DD")
    p.add_argument(
        "--last-week",
        action="store_true",
        help="使用上一週（週一至週日），預設為本週。",
    )

    # 郵件來源（本地）。
    p.add_argument("--mbox", help="本地 .mbox 發件箱檔")
    p.add_argument("--eml", help="單一 .eml 檔，或內含多個 .eml 的資料夾")

    # 郵件來源（IMAP）。
    p.add_argument("--imap-host", help="IMAP 主機")
    p.add_argument("--imap-user", help="IMAP 帳號")
    p.add_argument("--imap-password", help="IMAP 密碼（建議改用環境變數或設定檔）")
    p.add_argument("--imap-port", type=int, default=993, help="IMAP 連接埠（預設 993）")
    p.add_argument("--no-ssl", action="store_true", help="IMAP 不使用 SSL")
    p.add_argument("--sent-folder", help="寄件備份資料夾名稱（不指定則自動偵測）")

    # 額外輸入。
    p.add_argument(
        "--other",
        action="append",
        default=[],
        metavar="事項",
        help="其他事項，可重複指定多個。",
    )
    p.add_argument(
        "--interactive",
        action="store_true",
        help="互動模式：逐項詢問缺少的資訊與其他事項。",
    )
    p.add_argument(
        "--include-empty-days",
        action="store_true",
        help="保留沒有任何工作內容的日期。",
    )

    # 輸出。
    p.add_argument(
        "-f",
        "--format",
        default="text",
        choices=sorted(RENDERERS.keys()),
        help="輸出格式（預設 text）。",
    )
    p.add_argument("-o", "--output", help="輸出檔路徑（不指定則印到畫面）。")
    p.add_argument("--version", action="version", version=f"weekly_report {__version__}")
    return p


def _resolve_config(args: argparse.Namespace) -> ReportConfig:
    if args.config:
        config = load_config(args.config)
    else:
        config = config_from_dict({})

    if args.name:
        config.name = args.name
    if args.department:
        config.department = args.department
    if args.title:
        config.title = args.title

    if args.start and args.end:
        config.period_start = parse_date(args.start)
        config.period_end = parse_date(args.end)
    elif args.last_week:
        config.period_start, config.period_end = week_range(offset_weeks=-1)

    if args.other:
        config.other_matters = list(config.other_matters) + list(args.other)

    if args.interactive:
        _interactive_fill(config)

    return config


def _interactive_fill(config: ReportConfig) -> None:
    print("=== 互動輸入（直接按 Enter 可略過）===", file=sys.stderr)
    if not config.name:
        config.name = input("姓名：").strip() or config.name
    if not config.department:
        config.department = input("部門：").strip() or config.department

    print("輸入其他事項，一行一項，空行結束：", file=sys.stderr)
    while True:
        try:
            line = input("- ").strip()
        except EOFError:
            break
        if not line:
            break
        config.other_matters.append(line)


def _emails_for(
    args: argparse.Namespace,
    config: ReportConfig,
) -> List[EmailMessage]:
    start, end = config.period_start, config.period_end

    # 1) 命令列指定的本地檔。
    if args.mbox:
        return load_from_file(args.mbox, start, end)
    if args.eml:
        return load_from_file(args.eml, start, end)

    # 2) 命令列指定的 IMAP。
    if args.imap_host and args.imap_user:
        password = args.imap_password or os.environ.get("WEEKLY_REPORT_IMAP_PASSWORD")
        if not password:
            password = getpass.getpass("IMAP 密碼：")
        return load_from_imap(
            host=args.imap_host,
            username=args.imap_user,
            password=password,
            period_start=start,
            period_end=end,
            port=args.imap_port,
            use_ssl=not args.no_ssl,
            sent_folder=args.sent_folder,
        )

    # 3) 設定檔中的 email 區塊。
    if args.config:
        settings = email_settings(args.config)
        return _emails_from_settings(settings, start, end)

    # 4) 沒有任何郵件來源 → 只用手動／互動輸入。
    return []


def _emails_from_settings(settings: dict, start: date, end: date) -> List[EmailMessage]:
    if not settings:
        return []
    source = str(settings.get("source", "")).lower()

    if source in ("mbox", "eml", "file") or settings.get("path"):
        return load_from_file(settings["path"], start, end)

    if source == "imap" or settings.get("host"):
        password = (
            settings.get("password")
            or os.environ.get(settings.get("password_env", ""))
            or os.environ.get("WEEKLY_REPORT_IMAP_PASSWORD")
        )
        if not password:
            password = getpass.getpass("IMAP 密碼：")
        return load_from_imap(
            host=settings["host"],
            username=settings["username"],
            password=password,
            period_start=start,
            period_end=end,
            port=int(settings.get("port", 993)),
            use_ssl=bool(settings.get("ssl", True)),
            sent_folder=settings.get("sent_folder"),
        )
    return []


def main(argv: Optional[List[str]] = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)

    try:
        config = _resolve_config(args)
        if not config.name:
            parser.error("缺少姓名，請用 --name 或設定檔提供，或加上 --interactive。")

        emails = _emails_for(args, config)
        report = build_report(
            config, emails, include_empty_days=args.include_empty_days
        )
        output = render(report, args.format)
    except (FileNotFoundError, ValueError, RuntimeError) as exc:
        print(f"錯誤：{exc}", file=sys.stderr)
        return 1

    if args.output:
        with open(args.output, "w", encoding="utf-8") as fh:
            fh.write(output)
        print(
            f"已輸出週報到 {args.output}"
            f"（共 {report.total_items} 條工作內容，{len(report.work_days)} 天）。",
            file=sys.stderr,
        )
    else:
        print(output)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
