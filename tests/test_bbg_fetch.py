"""bbg_fetch 純 Python 輔助函式（不需 Excel）。"""

from lme_daily.bbg_fetch import (
    _as_2d_tuple,
    _read_range,
    fetch_bbg_3m_prompt_date,
    format_3m_date_for_vba,
    normalize_bbg_cell,
    parse_bbg_prompt_date_value,
    parse_bdp_formula,
    parse_excel_display_text,
    value_from_stored_number,
)


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


def test_parse_excel_display_text_thousands_separator():
    value = parse_excel_display_text("16,554.09")
    assert value == 16554.09
    assert isinstance(value, float)


def test_parse_excel_display_text_na_invalid_security():
    assert parse_excel_display_text("#N/A Invalid Security") == "N/A"


def test_parse_excel_display_text_empty():
    assert parse_excel_display_text("") is None


def test_normalize_bbg_cell_after_display_parse_keeps_na():
    parsed = parse_excel_display_text("#N/A Invalid Security")
    assert parsed == "N/A"
    assert normalize_bbg_cell(parsed) == "N/A"
    parsed_prefix = parse_excel_display_text("N/A Field Not Applicable")
    assert parsed_prefix == "N/A"
    assert normalize_bbg_cell(parsed_prefix) == "N/A"
    parsed_num = parse_excel_display_text("16,554.09")
    assert normalize_bbg_cell(parsed_num) == 16554.09


def test_value_from_stored_number_rounds_like_excel_display():
    assert value_from_stored_number(16554.08984375) == 16554.09
    assert value_from_stored_number(1838.129999) == 1838.13
    assert value_from_stored_number("N/A") == "N/A"
    assert value_from_stored_number(None) is None


def test_read_range_uses_cell_text_not_value2():
    from unittest.mock import MagicMock

    texts = (("16,554.09", "#N/A Invalid Security", ""),)

    def cells(r, c):
        cell = MagicMock()
        cell.Text = texts[r - 1][c - 1]
        cell.Value2 = 16554.08984375
        return cell

    rng = MagicMock()
    rng.Rows.Count = 1
    rng.Columns.Count = 3
    rng.Cells.side_effect = cells
    rng.NumberFormat = (("0.00", "General", "General"),)
    workbook = MagicMock()
    workbook.Worksheets.return_value.Range.return_value = rng

    values, formats = _read_range(workbook, "Promt date", "B3:D3")
    assert values == ((16554.09, "N/A", None),)
    assert formats[0][0] == "0.00"
    assert normalize_bbg_cell(values[0][1]) == "N/A"


def test_overlay_bdp_grid_replaces_formulas():
    from lme_daily.bbg_blpapi import _overlay_values

    grid = [["CA", '=BDP("LMCADS03 Comdty","PX_LAST")']]
    out = _overlay_values(grid, {("LMCADS03 Comdty", "PX_LAST"): 9001.5}, None)  # type: ignore[arg-type]
    assert out == (("CA", 9001.5),)


def _bbg_workbook_mock(*, prompt_value2=46345.0, prompt_text="11/19/2026"):
    from unittest.mock import MagicMock

    texts = (
        ("Metal", "N/A Field Not Applicable"),
        ("CA", "14,234.50"),
    )

    def cells(r, c):
        cell = MagicMock()
        cell.Text = texts[r - 1][c - 1]
        return cell

    workbook = MagicMock()
    workbook.Close = MagicMock()
    rng = MagicMock()
    rng.Rows.Count = 2
    rng.Columns.Count = 2
    rng.Cells.side_effect = cells
    rng.NumberFormat = (("General", "General"), ("General", "0.00"))
    prompt_cell = MagicMock()
    prompt_cell.Value2 = prompt_value2
    prompt_cell.Text = prompt_text

    def range_for(addr):
        if isinstance(addr, str) and ":" in addr:
            return rng
        return prompt_cell

    sheet = MagicMock()
    sheet.Range.side_effect = range_for
    workbook.Worksheets.return_value = sheet
    return workbook


def test_bbg_fetch_opens_when_workbook_not_already_open(tmp_path, monkeypatch):
    from contextlib import contextmanager
    from unittest.mock import MagicMock

    from lme_daily.bbg_fetch import fetch_bloomberg_snapshot
    from lme_daily.config import load_config
    from tests.test_excel_reuse import FakeComError, _write_config

    config = load_config(_write_config(tmp_path))
    workbook = _bbg_workbook_mock()
    slept: list[float] = []

    class Workbooks:
        def __init__(self) -> None:
            self.Open = MagicMock(return_value=workbook)

        def __call__(self, name: str):
            assert name == config.paths.bbg_workbook_name
            raise FakeComError()

    app = MagicMock()
    app.Workbooks = Workbooks()
    app.CalculationState = 0

    @contextmanager
    def fake_excel_app(**_kwargs):
        yield app

    monkeypatch.setattr("lme_daily.bbg_fetch.excel_app", fake_excel_app)
    monkeypatch.setattr("lme_daily.bbg_fetch.time.sleep", lambda seconds: slept.append(seconds))

    values, _formats = fetch_bloomberg_snapshot(config)
    assert values[1][1] == 14234.50
    app.Workbooks.Open.assert_called_once()
    kwargs = app.Workbooks.Open.call_args.kwargs
    assert kwargs.get("UpdateLinks") == 0
    assert kwargs.get("IgnoreReadOnlyRecommended") is True
    assert slept == [config.bloomberg.refresh_wait_seconds]
    workbook.Close.assert_not_called()
    workbook.RefreshAll.assert_called_once()


def test_bbg_fetch_reuses_already_open_workbook_without_open_or_close(tmp_path, monkeypatch):
    from contextlib import contextmanager
    from unittest.mock import MagicMock

    from lme_daily.bbg_fetch import fetch_bloomberg_snapshot
    from lme_daily.config import load_config
    from tests.test_excel_reuse import _write_config

    config = load_config(_write_config(tmp_path))
    workbook = _bbg_workbook_mock()
    slept: list[float] = []

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
    monkeypatch.setattr("lme_daily.bbg_fetch.time.sleep", lambda seconds: slept.append(seconds))

    values, _formats = fetch_bloomberg_snapshot(config)
    assert values[0][1] == "N/A"
    assert values[1][1] == 14234.50
    workbook.RefreshAll.assert_called_once()
    workbook.Close.assert_not_called()
    workbook.Activate.assert_not_called()
    app.Workbooks.Open.assert_not_called()
    assert slept == [config.bloomberg.refresh_wait_seconds]


def test_bbg_workbook_not_open_error_class_still_exists():
    from lme_daily.exceptions import BbgWorkbookNotOpenError

    assert issubclass(BbgWorkbookNotOpenError, Exception)


def test_parse_bbg_prompt_date_value_excel_serial():
    assert parse_bbg_prompt_date_value(46345) == "20261119"
    assert parse_bbg_prompt_date_value(46345.0) == "20261119"
    assert isinstance(parse_bbg_prompt_date_value(46345), str)


def test_parse_bbg_prompt_date_value_na():
    assert parse_bbg_prompt_date_value("N/A") is None
    assert parse_bbg_prompt_date_value("#N/A Invalid Security") is None
    assert parse_bbg_prompt_date_value("N/A Field Not Applicable") is None


def test_parse_bbg_prompt_date_value_blank():
    assert parse_bbg_prompt_date_value("") is None
    assert parse_bbg_prompt_date_value("   ") is None
    assert parse_bbg_prompt_date_value(None) is None


def test_parse_bbg_prompt_date_value_text_and_datetime():
    from datetime import date, datetime

    assert parse_bbg_prompt_date_value("20261119") == "20261119"
    assert parse_bbg_prompt_date_value("2026/11/19") == "20261119"
    assert parse_bbg_prompt_date_value(datetime(2026, 11, 19)) == "20261119"
    assert parse_bbg_prompt_date_value(date(2026, 12, 1)) == "20261201"


def test_format_3m_date_for_vba():
    assert format_3m_date_for_vba("20261119", "%Y/%m/%d") == "2026/11/19"
    assert format_3m_date_for_vba("20261119", "%Y%m%d") == "20261119"


def test_fetch_bbg_3m_prompt_date_from_value2_serial(tmp_path):
    from unittest.mock import MagicMock

    from lme_daily.config import load_config
    from tests.test_excel_reuse import _write_config

    config = load_config(_write_config(tmp_path))
    assert config.bloomberg.prompt_date_cell == "B4"
    workbook = _bbg_workbook_mock(prompt_value2=46345.0, prompt_text="n/a")
    app = MagicMock()
    app.Workbooks = MagicMock(return_value=workbook)
    ymd = fetch_bbg_3m_prompt_date(app, config, fallback="19990101")
    assert ymd == "20261119"
    workbook.Worksheets.assert_called_with(config.bloomberg.bbg_sheet_name)
    workbook.Worksheets.return_value.Range.assert_called_with("B4")


def test_fetch_bbg_3m_prompt_date_na_falls_back(tmp_path, caplog):
    from unittest.mock import MagicMock

    from lme_daily.config import load_config
    from tests.test_excel_reuse import _write_config

    config = load_config(_write_config(tmp_path))
    workbook = _bbg_workbook_mock(prompt_value2="N/A", prompt_text="#N/A Invalid Security")
    app = MagicMock()
    app.Workbooks = MagicMock(return_value=workbook)
    with caplog.at_level("WARNING"):
        ymd = fetch_bbg_3m_prompt_date(app, config, fallback="20261119")
    assert ymd == "20261119"
    assert "無法解析" in caplog.text


def test_fetch_bbg_3m_prompt_date_blank_falls_back(tmp_path, caplog):
    from unittest.mock import MagicMock

    from lme_daily.config import load_config
    from tests.test_excel_reuse import _write_config

    config = load_config(_write_config(tmp_path))
    workbook = _bbg_workbook_mock(prompt_value2=None, prompt_text="")
    app = MagicMock()
    app.Workbooks = MagicMock(return_value=workbook)
    with caplog.at_level("WARNING"):
        ymd = fetch_bbg_3m_prompt_date(app, config, fallback="20260818")
    assert ymd == "20260818"
    assert "無法解析" in caplog.text


def test_fetch_bbg_3m_prompt_date_uses_text_when_value2_unusable(tmp_path):
    from unittest.mock import MagicMock

    from lme_daily.config import load_config
    from tests.test_excel_reuse import _write_config

    config = load_config(_write_config(tmp_path))
    workbook = _bbg_workbook_mock(prompt_value2=None, prompt_text="2026/12/01")
    app = MagicMock()
    app.Workbooks = MagicMock(return_value=workbook)
    assert fetch_bbg_3m_prompt_date(app, config, fallback="19990101") == "20261201"
