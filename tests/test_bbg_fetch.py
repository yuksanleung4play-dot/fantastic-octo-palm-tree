"""bbg_fetch 純 Python 輔助函式（不需 Excel）。"""

from lme_daily.bbg_fetch import _as_2d_tuple


def test_as_2d_tuple_none():
    assert _as_2d_tuple(None) == ()


def test_as_2d_tuple_scalar():
    assert _as_2d_tuple(12.5) == ((12.5,),)


def test_as_2d_tuple_single_row():
    assert _as_2d_tuple((1, 2, 3)) == ((1, 2, 3),)


def test_as_2d_tuple_matrix():
    raw = ((1, 2), (3, 4))
    assert _as_2d_tuple(raw) == raw


def test_as_2d_tuple_formats_as_str():
    assert _as_2d_tuple(("0.0", "General"), as_str=True) == (("0.0", "General"),)
