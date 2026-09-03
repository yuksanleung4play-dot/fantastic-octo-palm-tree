"""calc_lme_dates / 3M 規則單元測試（不需 Excel）。"""

from __future__ import annotations

from datetime import date

import pytest

from lme_daily.dates import (
    calc_lme_dates,
    is_settlement_business_day,
    lme_3m_date,
    previous_settlement_business_day,
)
from lme_daily.exceptions import DateCalcError


def test_weekend_is_not_sbd():
    assert is_settlement_business_day(date(2026, 8, 19)) is True  # Wed
    assert is_settlement_business_day(date(2026, 8, 22)) is False  # Sat
    assert is_settlement_business_day(date(2026, 8, 23)) is False  # Sun


def test_holiday_is_not_sbd():
    holidays = {date(2026, 8, 19)}
    assert is_settlement_business_day(date(2026, 8, 19), holidays) is False
    assert is_settlement_business_day(date(2026, 8, 18), holidays) is True


def test_previous_sbd_skips_weekend():
    # Monday 2026-08-17 → previous SBD is Friday 2026-08-14
    assert previous_settlement_business_day(date(2026, 8, 17)) == date(2026, 8, 14)


def test_previous_sbd_skips_holiday_and_weekend():
    # Tuesday after Monday holiday → Friday
    holidays = {date(2026, 8, 17)}  # Monday
    assert previous_settlement_business_day(date(2026, 8, 18), holidays) == date(2026, 8, 14)


def test_previous_sbd_from_midweek():
    assert previous_settlement_business_day(date(2026, 8, 19)) == date(2026, 8, 18)


def test_3m_same_weekday_when_already_sbd():
    # 2026-08-19 + 3m = 2026-11-19 (Thursday)
    assert lme_3m_date(date(2026, 8, 19)) == date(2026, 11, 19)


def test_3m_rolls_forward_within_month():
    # 2026-08-20 + 3m = 2026-11-20 (Friday) already SBD
    assert lme_3m_date(date(2026, 8, 20)) == date(2026, 11, 20)
    # 2026-08-21 + 3m = 2026-11-21 (Saturday) → 2026-11-23 Monday (same month)
    assert lme_3m_date(date(2026, 8, 21)) == date(2026, 11, 23)


def test_3m_rolls_back_when_forward_crosses_month():
    # 2024-03-30 + 3m = 2024-06-30 (Sunday)
    # forward → 2024-07-01 (month change) → back to 2024-06-28 Friday
    assert lme_3m_date(date(2024, 3, 30)) == date(2024, 6, 28)


def test_3m_rolls_back_from_saturday_month_end():
    # 2024-03-29 + 3m = 2024-06-29 (Saturday)
    # forward → 2024-07-01 → back to 2024-06-28
    assert lme_3m_date(date(2024, 3, 29)) == date(2024, 6, 28)


def test_3m_holiday_rolls_forward_same_month():
    holidays = {date(2024, 12, 25), date(2024, 12, 26)}
    # 2024-09-25 + 3m = 2024-12-25 (Wed, holiday) → 2024-12-27 Friday
    assert lme_3m_date(date(2024, 9, 25), holidays) == date(2024, 12, 27)


def test_3m_month_end_clamp():
    # 2026-01-31 + 3m → 2026-04-30 (relativedelta clamp); 30 Apr 2026 is Thursday
    assert lme_3m_date(date(2026, 1, 31)) == date(2026, 4, 30)


def test_calc_lme_dates_returns_formatted_strings():
    prev, three_m = calc_lme_dates(date(2026, 8, 19), date_format="%Y/%m/%d")
    assert prev == "2026/08/18"
    assert three_m == "2026/11/19"


def test_calc_lme_dates_yyyymmdd_for_inputbox():
    prev, three_m = calc_lme_dates(date(2026, 8, 19), date_format="%Y%m%d")
    assert prev == "20260818"
    assert three_m == "20261119"


def test_calc_lme_dates_custom_format():
    prev, three_m = calc_lme_dates(date(2026, 8, 19), date_format="%d-%b-%Y")
    assert prev == "18-Aug-2026"
    assert three_m == "19-Nov-2026"


def test_calc_lme_dates_rejects_non_date():
    with pytest.raises(DateCalcError):
        calc_lme_dates("2026-08-19")  # type: ignore[arg-type]
