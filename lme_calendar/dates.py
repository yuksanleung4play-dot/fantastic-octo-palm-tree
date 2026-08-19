"""LME Settlement Business Day, Cash Date, and 3M Date calculations.

Cash Date follows the spec: the 2nd Settlement Business Day after *today* (T+2).

3M Date follows LME Rulebook Part 3, Trading Regulation 8.4.1 / 8.4.2, which is
the precise form of the "roll forward, unless that would change the month" rule:

  8.4.1  If the unadjusted date is not a Settlement Business Day, move to the
         next Settlement Business Day, EXCEPT:
           (a) Saturday, and the preceding Friday is a Settlement Business Day
               → previous Settlement Business Day (typically Friday)
           (b) Good Friday → previous Settlement Business Day
           (c) Christmas Day falling Tue–Fri → previous Settlement Business Day
           (d) Exchange-declared non-Business Day → Exchange discretion;
               market convention for the rolling 3M is next SBD, then 8.4.2
  8.4.2  If the 3M prompt would land in the 4th calendar month after the trade
         month, use the last Settlement Business Day of the 3rd calendar month.

This is why Trade Date 2026-08-07 (Fri) → 3M target 2026-11-07 (Sat) → 2026-11-06
rather than the naive following Monday 2026-11-09.
"""

from __future__ import annotations

import warnings
from datetime import date, timedelta
from typing import Iterable, Mapping

from dateutil.relativedelta import relativedelta

from lme_calendar.holidays import (
    HolidayCalendar,
    HolidayCoverageWarning,
    as_calendar,
)

WEEKDAY_NAMES = (
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday",
)


def weekday_name(day: date) -> str:
    return WEEKDAY_NAMES[day.weekday()]


def _easter_gregorian(year: int) -> date:
    """Anonymous Gregorian algorithm — used to recognise Good Friday."""
    a = year % 19
    b, c = divmod(year, 100)
    d, e = divmod(b, 4)
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = divmod(c, 4)
    ll = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * ll) // 451
    month, day = divmod(h + ll - 7 * m + 114, 31)
    return date(year, month, day + 1)


def is_good_friday(day: date) -> bool:
    return day == _easter_gregorian(day.year) - timedelta(days=2)


def is_christmas_day(day: date) -> bool:
    return day.month == 12 and day.day == 25


def is_weekend(day: date) -> bool:
    return day.weekday() >= 5


def is_business_day(day: date, holidays: HolidayCalendar | Iterable[date] | Mapping[date, object]) -> bool:
    """True if *day* is an LME Settlement Business Day (valid USD/GBP prompt).

    Settlement Business Day = weekday AND not a UK bank holiday AND not a US
    bank holiday. Either region's holiday makes the date invalid as a prompt.
    """
    calendar = as_calendar(holidays)
    if is_weekend(day):
        return False
    return not calendar.is_holiday(day)


def _check_coverage(day: date, calendar: HolidayCalendar, context: str) -> None:
    if day < calendar.coverage_start or day > calendar.coverage_end:
        warnings.warn(
            f"{context} {day.isoformat()} is outside holiday coverage "
            f"{calendar.coverage_start.isoformat()} .. {calendar.coverage_end.isoformat()}. "
            "Update the holiday list.",
            HolidayCoverageWarning,
            stacklevel=3,
        )


def _next_business_day(day: date, calendar: HolidayCalendar, *, context: str) -> date:
    cursor = day
    for _ in range(400):
        cursor += timedelta(days=1)
        _check_coverage(cursor, calendar, context)
        if is_business_day(cursor, calendar):
            return cursor
    raise RuntimeError(f"Could not find a Settlement Business Day after {day.isoformat()}")


def _previous_business_day(day: date, calendar: HolidayCalendar, *, context: str) -> date:
    cursor = day
    for _ in range(400):
        cursor -= timedelta(days=1)
        _check_coverage(cursor, calendar, context)
        if is_business_day(cursor, calendar):
            return cursor
    raise RuntimeError(f"Could not find a Settlement Business Day before {day.isoformat()}")


def calc_cash_date(
    today: date, holidays: HolidayCalendar | Iterable[date] | Mapping[date, object]
) -> date:
    """Cash prompt = 2nd Settlement Business Day after *today* (T+2, not T+2 calendar days)."""
    calendar = as_calendar(holidays)
    _check_coverage(today, calendar, "Trade Date")
    cursor = today
    found = 0
    while found < 2:
        cursor += timedelta(days=1)
        _check_coverage(cursor, calendar, "Cash Date")
        if is_business_day(cursor, calendar):
            found += 1
    return cursor


def _third_and_fourth_month(trade_date: date) -> tuple[date, date]:
    """First day of the 3rd / 4th calendar month after the trade month."""
    trade_month = date(trade_date.year, trade_date.month, 1)
    third = trade_month + relativedelta(months=3)
    fourth = trade_month + relativedelta(months=4)
    return third, fourth


def _last_business_day_of_month(month_start: date, calendar: HolidayCalendar, *, context: str) -> date:
    next_month = month_start + relativedelta(months=1)
    return _previous_business_day(next_month, calendar, context=context)


def calc_3m_date(
    today: date, holidays: HolidayCalendar | Iterable[date] | Mapping[date, object]
) -> date:
    """3-month prompt date from *today*, with LME 8.4.1 / 8.4.2 adjustments."""
    calendar = as_calendar(holidays)
    _check_coverage(today, calendar, "Trade Date")
    target = today + relativedelta(months=3)
    _check_coverage(target, calendar, "3M Date (unadjusted)")

    if is_business_day(target, calendar):
        result = target
    elif target.weekday() == 5 and is_business_day(target - timedelta(days=1), calendar):
        # 8.4.1(a) Saturday → preceding Friday when that Friday is a SBD
        result = target - timedelta(days=1)
    elif is_good_friday(target):
        result = _previous_business_day(target, calendar, context="3M Date")
    elif is_christmas_day(target) and target.weekday() in {1, 2, 3, 4}:
        result = _previous_business_day(target, calendar, context="3M Date")
    else:
        result = _next_business_day(target, calendar, context="3M Date")

    _third, fourth = _third_and_fourth_month(today)
    if (result.year, result.month) >= (fourth.year, fourth.month):
        result = _last_business_day_of_month(_third, calendar, context="3M Date (month cap)")

    _check_coverage(result, calendar, "3M Date")
    return result


def generate_trading_calendar(
    start_date: date,
    months_forward: int,
    holidays: HolidayCalendar | Iterable[date] | Mapping[date, object],
) -> list[date]:
    """Valid LME Settlement Business Days from *start_date* through +*months_forward* months."""
    if months_forward < 1:
        raise ValueError("months_forward must be >= 1")
    calendar = as_calendar(holidays)
    end_date = start_date + relativedelta(months=months_forward)
    _check_coverage(start_date, calendar, "Calendar start")
    _check_coverage(end_date, calendar, "Calendar end")

    days: list[date] = []
    cursor = start_date
    while cursor <= end_date:
        if is_business_day(cursor, calendar):
            days.append(cursor)
        cursor += timedelta(days=1)
    return days


def iter_calendar_days(start_date: date, months_forward: int) -> list[date]:
    """Every calendar day in the rolling window (including weekends/holidays)."""
    end_date = start_date + relativedelta(months=months_forward)
    days: list[date] = []
    cursor = start_date
    while cursor <= end_date:
        days.append(cursor)
        cursor += timedelta(days=1)
    return days
