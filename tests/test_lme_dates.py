from datetime import date
from pathlib import Path

import pytest

from lme_calendar.dates import calc_3m_date, calc_cash_date, generate_trading_calendar, is_business_day
from lme_calendar.holidays import HolidayFormatError, HolidayFileError, load_holiday_csv

HOLIDAYS_PATH = Path(__file__).resolve().parents[1] / "data" / "holidays.csv"


@pytest.fixture(scope="module")
def holidays():
    return load_holiday_csv(HOLIDAYS_PATH)


class TestValidationAnchor:
    """Spec reference: Trade 2026-08-07 → Cash 2026-08-11, 3M 2026-11-06."""

    TRADE = date(2026, 8, 7)

    def test_cash_date(self, holidays):
        assert calc_cash_date(self.TRADE, holidays) == date(2026, 8, 11)

    def test_three_m_date_rolls_back_from_saturday(self, holidays):
        # Unadjusted 2026-11-07 is Saturday; 8.4.1(a) → Friday 2026-11-06,
        # not the following Monday 2026-11-09.
        assert calc_3m_date(self.TRADE, holidays) == date(2026, 11, 6)

    def test_three_m_does_not_follow_to_monday(self, holidays):
        assert calc_3m_date(self.TRADE, holidays) != date(2026, 11, 9)


class TestBusinessDay:
    def test_weekend_is_not_business_day(self, holidays):
        assert is_business_day(date(2026, 8, 8), holidays) is False  # Saturday
        assert is_business_day(date(2026, 8, 9), holidays) is False  # Sunday

    def test_weekday_is_business_day(self, holidays):
        assert is_business_day(date(2026, 8, 10), holidays) is True

    def test_uk_holiday_is_not_business_day(self, holidays):
        assert is_business_day(date(2026, 8, 31), holidays) is False  # Summer BH

    def test_us_holiday_is_not_business_day(self, holidays):
        assert is_business_day(date(2026, 11, 11), holidays) is False  # Veterans Day

    def test_jpy_only_day_remains_valid_usd_prompt(self, holidays):
        # 2026-08-11 is a Japanese holiday in the LME notice (Mountain Day) but
        # is still a valid USD/GBP prompt. It must NOT be in the UK/US file.
        assert is_business_day(date(2026, 8, 11), holidays) is True


class TestCashDate:
    def test_friday_skips_weekend(self, holidays):
        assert calc_cash_date(date(2026, 8, 7), holidays) == date(2026, 8, 11)

    def test_wednesday_is_t_plus_2(self, holidays):
        assert calc_cash_date(date(2026, 8, 19), holidays) == date(2026, 8, 21)

    def test_before_uk_holiday(self, holidays):
        # Friday 2026-08-28 → Mon 31 is UK holiday, so T+1=Tue 1 Sep, T+2=Wed 2 Sep
        assert calc_cash_date(date(2026, 8, 28), holidays) == date(2026, 9, 2)


class TestThreeMonthDate:
    def test_unadjusted_weekday(self, holidays):
        # 2026-08-10 + 3 months = 2026-11-10 (Tuesday, not a holiday)
        assert calc_3m_date(date(2026, 8, 10), holidays) == date(2026, 11, 10)

    def test_saturday_preceding_friday(self, holidays):
        assert calc_3m_date(date(2026, 8, 7), holidays) == date(2026, 11, 6)

    def test_sunday_follows_to_monday(self, holidays):
        # 2026-08-09 + 3 months = 2026-11-09 (Monday, not a holiday)
        assert calc_3m_date(date(2026, 8, 9), holidays) == date(2026, 11, 9)

    def test_us_holiday_follows(self, holidays):
        # 2026-08-11 + 3 months = 2026-11-11 Veterans Day (Wednesday)
        # Not Sat / Good Friday / Christmas → next SBD 2026-11-12
        assert calc_3m_date(date(2026, 8, 11), holidays) == date(2026, 11, 12)

    def test_christmas_friday_goes_preceding(self, holidays):
        # 2026-09-25 + 3 months = 2026-12-25 Friday Christmas → 8.4.1(c) → Dec 24
        assert calc_3m_date(date(2026, 9, 25), holidays) == date(2026, 12, 24)

    def test_good_friday_goes_preceding(self, holidays):
        # 2026-12-26 + 3 months = 2027-03-26 Good Friday → preceding SBD 2027-03-25
        assert calc_3m_date(date(2026, 12, 26), holidays) == date(2027, 3, 25)

    def test_month_cap_when_following_crosses_into_fourth_month(self, holidays):
        # 2026-11-30 + 3 months = 2027-02-28 (Sunday) → following would be 2027-03-01,
        # which is the 4th month after November. 8.4.2 → last SBD of February.
        result = calc_3m_date(date(2026, 11, 30), holidays)
        assert result == date(2027, 2, 26)
        assert result.month == 2


class TestTradingCalendar:
    def test_window_contains_only_valid_days(self, holidays):
        days = generate_trading_calendar(date(2026, 8, 19), 27, holidays)
        assert days[0] == date(2026, 8, 19)
        assert all(is_business_day(day, holidays) for day in days)
        assert date(2026, 8, 31) not in days  # UK holiday
        assert date(2026, 11, 11) not in days  # US holiday
        assert days[-1] <= date(2028, 11, 19)

    def test_rejects_non_positive_months(self, holidays):
        with pytest.raises(ValueError):
            generate_trading_calendar(date(2026, 8, 19), 0, holidays)


class TestHolidayFile:
    def test_missing_file_raises_clear_error(self, tmp_path):
        missing = tmp_path / "nope.csv"
        with pytest.raises(HolidayFileError, match="not found"):
            load_holiday_csv(missing)

    def test_wrong_columns_raise(self, tmp_path):
        bad = tmp_path / "bad.csv"
        bad.write_text("foo,bar\n1,2\n", encoding="utf-8")
        with pytest.raises(HolidayFormatError, match="missing columns"):
            load_holiday_csv(bad)

    def test_empty_data_raises(self, tmp_path):
        empty = tmp_path / "empty.csv"
        empty.write_text("Date,Region,Holiday_Name\n", encoding="utf-8")
        with pytest.raises(HolidayFormatError, match="no data rows"):
            load_holiday_csv(empty)

    def test_invalid_region_raises(self, tmp_path):
        bad = tmp_path / "region.csv"
        bad.write_text("Date,Region,Holiday_Name\n2026-01-01,JP,New Year\n", encoding="utf-8")
        with pytest.raises(HolidayFormatError, match="Invalid Region"):
            load_holiday_csv(bad)

    def test_loads_shipped_file(self, holidays):
        assert holidays.coverage_start == date(2026, 1, 1)
        assert holidays.coverage_end == date(2029, 12, 26)
        assert holidays.is_uk_holiday(date(2026, 8, 31))
        assert holidays.is_us_holiday(date(2026, 11, 11))
        assert holidays.is_uk_holiday(date(2026, 12, 25))
        assert holidays.is_us_holiday(date(2026, 12, 25))
