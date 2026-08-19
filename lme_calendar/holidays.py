"""Load and validate the UK/US bank-holiday file used for LME prompt dates."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Iterable, Mapping

import pandas as pd

REQUIRED_COLUMNS = ("Date", "Region", "Holiday_Name")
VALID_REGIONS = {"UK", "US", "BOTH"}

HOLIDAY_CSV_FORMAT = """\
Required CSV columns (header row, exact names):
  Date,Region,Holiday_Name

Rules:
  - Date must be YYYY-MM-DD
  - Region must be UK, US, or Both
  - Holiday_Name is a short label (e.g. "Christmas Day")
  - One row per holiday. If both regions observe the same date, use Region=Both
    (or supply two rows, one UK and one US; they will be merged).
  - Do not invent dates. Prefer a Bloomberg CDR London export, or the latest
    LME notice "Banking Holidays, Non-Valid Prompt Dates and LME Business Days".

Example:
  Date,Region,Holiday_Name
  2026-08-31,UK,Summer bank holiday
  2026-11-11,US,Veterans Day
  2026-12-25,Both,Christmas Day
"""


class HolidayFileError(FileNotFoundError):
    """Holiday file is missing or unreadable."""


class HolidayFormatError(ValueError):
    """Holiday file is present but does not match the required schema."""


@dataclass(frozen=True)
class Holiday:
    date: date
    region: str  # UK / US / Both
    name: str


class HolidayCalendar:
    """Holiday lookup with UK/US region flags and a coverage window."""

    def __init__(self, holidays: Iterable[Holiday]):
        by_date: dict[date, Holiday] = {}
        for item in holidays:
            region = _normalize_region(item.region)
            existing = by_date.get(item.date)
            if existing is None:
                by_date[item.date] = Holiday(item.date, region, item.name)
                continue
            merged_region = _merge_region(existing.region, region)
            merged_name = existing.name
            if item.name and item.name not in existing.name:
                merged_name = f"{existing.name} / {item.name}"
            by_date[item.date] = Holiday(item.date, merged_region, merged_name)

        if not by_date:
            raise HolidayFormatError(
                "Holiday file contains no data rows.\n\n" + HOLIDAY_CSV_FORMAT
            )

        self._by_date = dict(sorted(by_date.items()))
        self.coverage_start = min(self._by_date)
        self.coverage_end = max(self._by_date)

    @property
    def holidays(self) -> list[Holiday]:
        return list(self._by_date.values())

    def get(self, day: date) -> Holiday | None:
        return self._by_date.get(day)

    def is_uk_holiday(self, day: date) -> bool:
        item = self._by_date.get(day)
        return bool(item and item.region in {"UK", "Both"})

    def is_us_holiday(self, day: date) -> bool:
        item = self._by_date.get(day)
        return bool(item and item.region in {"US", "Both"})

    def is_holiday(self, day: date) -> bool:
        return day in self._by_date

    def all_dates(self) -> set[date]:
        return set(self._by_date)

    def ensure_covers(self, day: date, context: str) -> None:
        """Warn (via exception subclass? no — return message) if *day* is outside coverage.

        Callers print warnings; this method only reports.
        """
        if day < self.coverage_start or day > self.coverage_end:
            raise HolidayCoverageWarning(
                f"{context} {day.isoformat()} falls outside the holiday-file coverage "
                f"{self.coverage_start.isoformat()} .. {self.coverage_end.isoformat()}. "
                "Update the holiday list before relying on this result."
            )


class HolidayCoverageWarning(UserWarning):
    """Computed date is outside the holiday file's min/max dates."""


def _normalize_region(region: str) -> str:
    value = (region or "").strip().title()
    if value.upper() == "BOTH":
        return "Both"
    if value.upper() not in VALID_REGIONS:
        raise HolidayFormatError(
            f"Invalid Region '{region}'. Must be UK, US, or Both.\n\n" + HOLIDAY_CSV_FORMAT
        )
    return value.upper() if value.upper() in {"UK", "US"} else "Both"


def _merge_region(left: str, right: str) -> str:
    if left == right:
        return left
    return "Both"


def _parse_iso_date(value: str, *, loc: str) -> date:
    raw = (value or "").strip()
    try:
        return datetime.strptime(raw, "%Y-%m-%d").date()
    except ValueError as exc:
        raise HolidayFormatError(
            f"Invalid Date '{value}' ({loc}). Use YYYY-MM-DD.\n\n" + HOLIDAY_CSV_FORMAT
        ) from exc


def load_holiday_csv(path: str | Path) -> HolidayCalendar:
    """Load holidays from CSV. Raises a clear error if the file is missing or invalid."""
    csv_path = Path(path)
    if not csv_path.is_file():
        raise HolidayFileError(
            f"Holiday list file not found: {csv_path}\n\n"
            "Pass --holidays PATH to an existing CSV. "
            "Do not run without a holiday file — dates would be silently wrong.\n\n"
            + HOLIDAY_CSV_FORMAT
        )

    try:
        frame = pd.read_csv(csv_path, dtype=str, encoding="utf-8-sig", keep_default_na=False)
    except (OSError, UnicodeDecodeError) as exc:
        raise HolidayFileError(f"Cannot read holiday file {csv_path}: {exc}") from exc
    except pd.errors.EmptyDataError as exc:
        raise HolidayFormatError(
            f"Holiday file {csv_path} has no header row.\n\n" + HOLIDAY_CSV_FORMAT
        ) from exc
    except pd.errors.ParserError as exc:
        raise HolidayFormatError(
            f"Holiday file {csv_path} could not be parsed as CSV: {exc}\n\n" + HOLIDAY_CSV_FORMAT
        ) from exc

    columns = [str(name).strip() for name in frame.columns]
    frame.columns = columns
    missing = [col for col in REQUIRED_COLUMNS if col not in columns]
    if missing:
        raise HolidayFormatError(
            f"Holiday file {csv_path} is missing columns {missing}. "
            f"Found: {columns}\n\n" + HOLIDAY_CSV_FORMAT
        )

    rows: list[Holiday] = []
    for offset, record in enumerate(frame.to_dict(orient="records"), start=2):
        raw_date = str(record.get("Date", "")).strip()
        raw_region = str(record.get("Region", "")).strip()
        raw_name = str(record.get("Holiday_Name", "")).strip()
        if not raw_date and not raw_region and not raw_name:
            continue
        if not raw_date or not raw_region:
            raise HolidayFormatError(
                f"Incomplete row {offset} in {csv_path}. "
                "Date and Region are required.\n\n" + HOLIDAY_CSV_FORMAT
            )
        rows.append(
            Holiday(
                date=_parse_iso_date(raw_date, loc=f"{csv_path}:{offset}"),
                region=raw_region,
                name=raw_name or "Unnamed holiday",
            )
        )

    return HolidayCalendar(rows)


def as_calendar(holidays: HolidayCalendar | Iterable[date] | Mapping[date, object]) -> HolidayCalendar:
    """Accept a HolidayCalendar, a set of dates, or a date->label mapping."""
    if isinstance(holidays, HolidayCalendar):
        return holidays
    if isinstance(holidays, Mapping):
        items = [
            Holiday(day, "Both", str(label) if label is not None else "Holiday")
            for day, label in holidays.items()
        ]
        return HolidayCalendar(items)
    items = [Holiday(day, "Both", "Holiday") for day in holidays]
    return HolidayCalendar(items)
