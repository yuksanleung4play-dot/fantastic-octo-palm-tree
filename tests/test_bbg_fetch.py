"""bbg_fetch 純 Python 輔助函式（不需 Excel）。"""

from lme_daily.bbg_fetch import _as_2d_tuple, normalize_bbg_cell, parse_bdp_formula


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


def test_parse_bdp_formula():
    assert parse_bdp_formula('=BDP("LMCADS03 Comdty","PX_LAST")') == (
        "LMCADS03 Comdty",
        "PX_LAST",
    )
    assert parse_bdp_formula('BDP("CA Comdty","PX_BID")') == ("CA Comdty", "PX_BID")
    assert parse_bdp_formula(9001.5) is None
    assert parse_bdp_formula("Copper") is None


def test_normalize_bbg_cell_blank_stays_blank():
    assert normalize_bbg_cell(None) is None
    assert normalize_bbg_cell("") == ""
    assert normalize_bbg_cell("   ") == "   "


def test_normalize_bbg_cell_na_prefix_becomes_na():
    assert normalize_bbg_cell("N/A Field Not Applicable") == "N/A"
    assert normalize_bbg_cell("N/A Requesting Data...") == "N/A"
    assert normalize_bbg_cell("n/a ****") == "N/A"


def test_normalize_bbg_cell_number_unchanged():
    assert normalize_bbg_cell(14234.50) == 14234.50
    assert normalize_bbg_cell("Copper") == "Copper"


def test_overlay_bdp_grid_replaces_formulas():
    from lme_daily.bbg_blpapi import _overlay_values

    grid = [["CA", '=BDP("LMCADS03 Comdty","PX_LAST")']]
    out = _overlay_values(grid, {("LMCADS03 Comdty", "PX_LAST"): 9001.5}, None)  # type: ignore[arg-type]
    assert out == (("CA", 9001.5),)
