"""LME 上日日期與 3M date 計算。

可獨立做單元測試，不依賴 Excel / Bloomberg。
公休日請透過 ``holiday_list`` 傳入（config / holidays.yaml）；未提供時僅排除週末。
"""

from __future__ import annotations

from collections.abc import Iterable
from datetime import date, timedelta

from dateutil.relativedelta import relativedelta

from lme_daily.exceptions import DateCalcError

_MAX_SCAN_DAYS = 45


def is_settlement_business_day(
    d: date,
    holiday_list: Iterable[date] | None = None,
) -> bool:
    """判斷 ``d`` 是否為 Settlement Business Day。

    簡化交易日曆：週一至週五，且不在 ``holiday_list`` 中。
    LME 官方公休日請補進 holiday_list，否則會被當成交易日。
    """
    if d.weekday() >= 5:  # Saturday=5, Sunday=6
        return False
    holidays = set(holiday_list or ())
    return d not in holidays


def previous_settlement_business_day(
    today: date,
    holiday_list: Iterable[date] | None = None,
) -> date:
    """回傳嚴格早於 ``today`` 的最近一個 Settlement Business Day（上日日期）。

    從 ``today - 1`` 開始往前掃，跳過週末與公休日。
    例如週一執行 → 上星期五；若週一為假期且週二執行 → 上星期五。
    """
    d = today - timedelta(days=1)
    for _ in range(_MAX_SCAN_DAYS):
        if is_settlement_business_day(d, holiday_list):
            return d
        d -= timedelta(days=1)
    raise DateCalcError(
        f"無法在 {today.isoformat()} 之前 {_MAX_SCAN_DAYS} 天內找到 Settlement Business Day，"
        "請檢查 holiday_list 是否過密。"
    )


def next_settlement_business_day(
    d: date,
    holiday_list: Iterable[date] | None = None,
    *,
    inclusive: bool = True,
) -> date:
    """回傳 ``d`` 當日（若 ``inclusive``）或之後的最近一個 Settlement Business Day。"""
    cursor = d if inclusive else d + timedelta(days=1)
    for _ in range(_MAX_SCAN_DAYS):
        if is_settlement_business_day(cursor, holiday_list):
            return cursor
        cursor += timedelta(days=1)
    raise DateCalcError(
        f"無法在 {d.isoformat()} 之後 {_MAX_SCAN_DAYS} 天內找到 Settlement Business Day。"
    )


def previous_settlement_business_day_on_or_before(
    d: date,
    holiday_list: Iterable[date] | None = None,
) -> date:
    """回傳 ``d`` 當日或之前的最近一個 Settlement Business Day。"""
    cursor = d
    for _ in range(_MAX_SCAN_DAYS):
        if is_settlement_business_day(cursor, holiday_list):
            return cursor
        cursor -= timedelta(days=1)
    raise DateCalcError(
        f"無法在 {d.isoformat()} 之前 {_MAX_SCAN_DAYS} 天內找到 Settlement Business Day。"
    )


def lme_3m_date(
    today: date,
    holiday_list: Iterable[date] | None = None,
) -> date:
    """由 ``today`` 計算 LME 3M prompt date。

    規則：
    1. 往後推 3 個自然月，取同一天（``dateutil.relativedelta``；月底自動夾住）。
    2. 若該日不是 Settlement Business Day：先順延到下一個有效交易日。
    3. 若順延導致月份改變（跨月），改為往前推到上一個有效交易日，留在原 3M 月份。
    """
    target = today + relativedelta(months=3)
    if is_settlement_business_day(target, holiday_list):
        return target

    rolled_forward = next_settlement_business_day(target, holiday_list, inclusive=True)
    if (rolled_forward.year, rolled_forward.month) != (target.year, target.month):
        return previous_settlement_business_day_on_or_before(
            target, holiday_list
        )
    return rolled_forward


def calc_lme_dates(
    today: date,
    holiday_list: Iterable[date] | None = None,
    date_format: str = "%Y/%m/%d",
) -> tuple[str, str]:
    """計算巨集所需的兩個日期字串。

    Parameters
    ----------
    today:
        執行日（通常為 ``date.today()``，可用 ``--as-of`` 覆寫）。
    holiday_list:
        LME / 倫敦公休日。``None`` 或空則只排除週末。
    date_format:
        ``strftime`` 格式，需符合 Excel / 巨集 InputBox 地區設定。

    Returns
    -------
    tuple[str, str]
        ``(上日日期, 3M date)`` 格式化字串。
    """
    if not isinstance(today, date):
        raise DateCalcError(f"today 必須是 datetime.date，收到 {type(today)!r}")

    prev = previous_settlement_business_day(today, holiday_list)
    three_m = lme_3m_date(today, holiday_list)
    try:
        return prev.strftime(date_format), three_m.strftime(date_format)
    except (ValueError, TypeError) as exc:
        raise DateCalcError(f"日期格式化失敗（date_format={date_format!r}）：{exc}") from exc
