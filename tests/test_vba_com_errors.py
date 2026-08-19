"""COM HRESULT helpers for VBA Application.Run fallback."""

from lme_daily.exceptions import ExcelComError
from lme_daily.vba_runner import com_error_codes, is_bad_param_count


class FakeComError(Exception):
    """Mimic pywintypes.com_error args layout from the user's log."""

    def __init__(self) -> None:
        super().__init__(
            -2147352567,
            "發生例外狀況。",
            (0, None, None, None, 0, -2147352562),
            None,
        )


def test_bad_param_count_from_user_log_shape():
    exc = FakeComError()
    assert -2147352562 in com_error_codes(exc)
    assert is_bad_param_count(exc) is True


def test_bad_param_count_via_excel_com_error_cause():
    try:
        raise ExcelComError("參數模式失敗") from FakeComError()
    except ExcelComError as exc:
        assert is_bad_param_count(exc) is True
        assert is_bad_param_count(exc.__cause__) is True


def test_unrelated_exception_is_not_bad_param_count():
    assert is_bad_param_count(RuntimeError("nope")) is False
