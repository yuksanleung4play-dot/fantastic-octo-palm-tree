#!/usr/bin/env python3
"""Daily entry point: rebuild the LME trading calendar Excel file.

Designed for Windows Task Scheduler / cron. Only the sheets Trading_Calendar,
Cash_3M_Daily, and Holiday_Reference are modified; any other sheet in the
workbook is left untouched.
"""

from __future__ import annotations

import argparse
import logging
import sys
from datetime import date, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lme_calendar.excel import update_workbook
from lme_calendar.holidays import HolidayFileError, HolidayFormatError, load_holiday_csv

DEFAULT_EXCEL = ROOT / "output" / "lme_prompt_calendar.xlsx"
DEFAULT_HOLIDAYS = ROOT / "data" / "holidays.csv"

VALIDATION_TRADE = date(2026, 8, 7)
VALIDATION_CASH = date(2026, 8, 11)
VALIDATION_3M = date(2026, 11, 6)


def _parse_date(value: str) -> date:
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError as exc:
        raise argparse.ArgumentTypeError(f"Invalid date '{value}', expected YYYY-MM-DD") from exc


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Update LME Trading Calendar + Cash/3M Date Excel workbook."
    )
    parser.add_argument(
        "--excel",
        type=Path,
        default=DEFAULT_EXCEL,
        help=f"Workbook path (default: {DEFAULT_EXCEL})",
    )
    parser.add_argument(
        "--holidays",
        type=Path,
        default=DEFAULT_HOLIDAYS,
        help="Holiday CSV path. Required file; the script will not invent holiday dates.",
    )
    parser.add_argument(
        "--as-of",
        dest="as_of",
        type=_parse_date,
        default=date.today(),
        help="Trade/snapshot date YYYY-MM-DD (default: today).",
    )
    parser.add_argument(
        "--months",
        type=int,
        default=27,
        help="Rolling calendar window in months (default: 27).",
    )
    parser.add_argument(
        "--snapshot-mode",
        choices=("overwrite", "append"),
        default="overwrite",
        help="overwrite = keep a single current row on Cash_3M_Daily; "
        "append = accumulate daily history (same-day rerun replaces that day's row).",
    )
    parser.add_argument(
        "--retain-expired",
        action="store_true",
        help="Keep Trading_Calendar rows that have rolled out of the 27-month window.",
    )
    parser.add_argument(
        "--log-file",
        type=Path,
        default=None,
        help="Optional log file. The same summary is always printed to stdout.",
    )
    return parser


def _configure_logging(log_file: Path | None) -> logging.Logger:
    logger = logging.getLogger("lme_calendar")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    stream = logging.StreamHandler(sys.stdout)
    stream.setFormatter(formatter)
    logger.addHandler(stream)
    if log_file is not None:
        log_file.parent.mkdir(parents=True, exist_ok=True)
        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    return logger


def _log_validation_anchor(logger: logging.Logger, calendar) -> None:
    """Always log the spec's reference case so operators can see the logic still holds."""
    from lme_calendar.dates import calc_3m_date, calc_cash_date

    cash = calc_cash_date(VALIDATION_TRADE, calendar)
    three_m = calc_3m_date(VALIDATION_TRADE, calendar)
    ok = cash == VALIDATION_CASH and three_m == VALIDATION_3M
    logger.info(
        "Validation anchor Trade=%s Cash=%s (expected %s) 3M=%s (expected %s) [%s]",
        VALIDATION_TRADE.isoformat(),
        cash.isoformat(),
        VALIDATION_CASH.isoformat(),
        three_m.isoformat(),
        VALIDATION_3M.isoformat(),
        "PASS" if ok else "FAIL",
    )
    if not ok:
        logger.error(
            "Validation anchor failed. Check holiday file and date logic before using this workbook."
        )


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    logger = _configure_logging(args.log_file)

    try:
        calendar = load_holiday_csv(args.holidays)
    except (HolidayFileError, HolidayFormatError) as exc:
        logger.error("%s", exc)
        return 2

    logger.info("Holiday file: %s (%s rows, coverage %s .. %s)",
                args.holidays, len(calendar.holidays),
                calendar.coverage_start.isoformat(), calendar.coverage_end.isoformat())

    _log_validation_anchor(logger, calendar)

    try:
        result = update_workbook(
            args.excel,
            calendar=calendar,
            as_of=args.as_of,
            months_forward=args.months,
            snapshot_mode=args.snapshot_mode,
            retain_expired=args.retain_expired,
        )
    except Exception as exc:  # noqa: BLE001 — CLI must not swallow the message
        logger.exception("Failed to update workbook: %s", exc)
        return 1

    logger.info(
        "Updated %s | Snapshot_Date=%s Cash_Date=%s 3M_Date=%s Days_Cash_to_3M=%s "
        "valid_trading_days=%s mode=%s",
        result["excel_path"],
        result["snapshot_date"].isoformat(),
        result["cash_date"].isoformat(),
        result["three_m_date"].isoformat(),
        result["days_cash_to_3m"],
        result["valid_trading_days"],
        result["snapshot_mode"],
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
