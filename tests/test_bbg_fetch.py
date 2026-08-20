"""bbg_fetch 純 Python 輔助函式（不需 Excel）。"""

import pytest

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


def test_bbg_fetch_raises_when_workbook_not_open(tmp_path, monkeypatch):
    from contextlib import contextmanager
    from unittest.mock import MagicMock

    from lme_daily.bbg_fetch import fetch_bloomberg_snapshot
    from lme_daily.config import load_config
    from lme_daily.exceptions import BbgWorkbookNotOpenError
    from tests.test_excel_reuse import FakeComError, _write_config

    config = load_config(_write_config(tmp_path))
    app = MagicMock()
    app.Workbooks.side_effect = FakeComError()
    app.Workbooks.Open.side_effect = AssertionError("Workbooks.Open must not be called")

    @contextmanager
    def fake_excel_app(**_kwargs):
        yield app

    monkeypatch.setattr("lme_daily.bbg_fetch.excel_app", fake_excel_app)
    with pytest.raises(BbgWorkbookNotOpenError, match="請先手動打開"):
        fetch_bloomberg_snapshot(config)
    app.Workbooks.Open.assert_not_called()


def test_bbg_fetch_reads_already_open_workbook_without_open_or_close(tmp_path, monkeypatch):
    from contextlib import contextmanager
    from unittest.mock import MagicMock

    from lme_daily.bbg_fetch import fetch_bloomberg_snapshot
    from lme_daily.config import load_config
    from tests.test_excel_reuse import _write_config

    config = load_config(_write_config(tmp_path))
    workbook = MagicMock()
    workbook.Close = MagicMock()
    rng = MagicMock()
    rng.Value2 = (("Metal", "N/A Field Not Applicable"), ("CA", 14234.50))
    rng.NumberFormat = (("General", "General"), ("General", "0.00"))
    workbook.Worksheets.return_value.Range.return_value = rng

    class Workbooks:
        def __init__(self) -> None:
            self.Open = MagicMock(side_effect=AssertionError("Workbooks.Open must not be called"))

        def __call__(self, name: str):
            assert name == config.paths.bbg_workbook_name
            return workbook

    app = MagicMock()
    app.Workbooks = Workbooks()
    app.CalculationState = 0

    @contextmanager
    def fake_excel_app(**_kwargs):
        yield app

    monkeypatch.setattr("lme_daily.bbg_fetch.excel_app", fake_excel_app)
    values, _formats = fetch_bloomberg_snapshot(config)
    assert values[0][1] == "N/A"
    assert values[1][1] == 14234.50
    workbook.RefreshAll.assert_called_once()
    workbook.Close.assert_not_called()
    workbook.Activate.assert_not_called()
    app.Workbooks.Open.assert_not_called()
