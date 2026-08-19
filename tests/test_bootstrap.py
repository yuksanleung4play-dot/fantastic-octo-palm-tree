"""bootstrap 套件檢查。"""

from lme_daily.bootstrap import _missing_packages


def test_core_packages_present_on_this_runner():
    assert _missing_packages(need_win32=False) == []
