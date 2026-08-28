"""P: → UNC rewrite and Excel attach diagnostics."""

from __future__ import annotations

import logging
from pathlib import Path
from unittest.mock import MagicMock

import pytest

from lme_daily.excel_com import (
    _acquire_running_excel,
    rewrite_workbook_p_drive_to_unc,
    run_p_drive_visibility_macro,
)
from lme_daily.exceptions import ExcelComError
from lme_daily.unc_paths import P_DRIVE_UNC_ROOT, rewrite_p_drive_to_unc


def test_rewrite_p_drive_to_unc_span_dat():
    src = r"TEXT;P:\Dealing Department - New\1. 交易部日常工作分類\2. 期貨\LME SPAN\lme.20260827.dat"
    out = rewrite_p_drive_to_unc(src)
    assert out.startswith("TEXT;" + P_DRIVE_UNC_ROOT)
    assert r"\Dealing Department - New\1. 交易部日常工作分類" in out
    assert "P:" not in out
    assert out == rewrite_p_drive_to_unc(out)


def test_rewrite_p_drive_forward_slash():
    src = "P:/Dealing Department - New/LME SPAN/lme.dat"
    out = rewrite_p_drive_to_unc(src)
    assert out.startswith(P_DRIVE_UNC_ROOT.replace("\\", "/"))
    assert not out.upper().startswith("P:")


def test_rewrite_leaves_unc_and_unrelated_text():
    unc = r"\\192.168.89.167\Dealing\Dealing Department - New\file.dat"
    assert rewrite_p_drive_to_unc(unc) == unc
    assert rewrite_p_drive_to_unc("no mapped drive here") == "no mapped drive here"


class _Coll:
    def __init__(self, items: list[object]) -> None:
        self._items = items

    @property
    def Count(self) -> int:
        return len(self._items)

    def Item(self, index: int) -> object:
        return self._items[index - 1]


class _QueryTable:
    def __init__(self, connection: str) -> None:
        self.Connection = connection
        self.CommandText = ""


class _Sheet:
    def __init__(self, qt: _QueryTable) -> None:
        self.QueryTables = _Coll([qt])
        self.ListObjects = _Coll([])


class _Workbook:
    def __init__(self, qt: _QueryTable) -> None:
        self.Worksheets = _Coll([_Sheet(qt)])
        self.Connections = _Coll([])

    @property
    def VBProject(self) -> object:
        raise RuntimeError("programmatic access to Visual Basic Project is not trusted")


def test_rewrite_workbook_querytable_connection():
    qt = _QueryTable(
        r"TEXT;P:\Dealing Department - New\LME SPAN\lme.20260827.dat"
    )
    changed = rewrite_workbook_p_drive_to_unc(_Workbook(qt))
    assert changed >= 1
    assert qt.Connection.startswith("TEXT;" + P_DRIVE_UNC_ROOT)
    assert "P:" not in qt.Connection


def test_run_p_drive_visibility_macro_uses_workbook_qualified_name(caplog: pytest.LogCaptureFixture):
    app = MagicMock()
    app.Run.return_value = "P: 磁碟機看不到！"
    workbook = MagicMock()
    workbook.Name = "早班_LME_reference_2024.xlsm"
    with caplog.at_level(logging.INFO):
        result = run_p_drive_visibility_macro(app, workbook)
    assert result == "P: 磁碟機看不到！"
    app.Run.assert_called_once_with("'早班_LME_reference_2024.xlsm'!TestPDriveVisible")
    assert "P: 磁碟機可見性測試結果：P: 磁碟機看不到！" in caplog.text


def test_run_p_drive_visibility_macro_missing_logs_warning(caplog: pytest.LogCaptureFixture):
    app = MagicMock()
    app.Run.side_effect = RuntimeError("unknown name")
    workbook = MagicMock()
    workbook.Name = "ref.xlsm"
    with caplog.at_level(logging.WARNING):
        assert run_p_drive_visibility_macro(app, workbook) is None
    assert "無法執行 TestPDriveVisible" in caplog.text


def test_acquire_running_excel_never_calls_dispatch():
    class Client:
        def GetActiveObject(self, name: str) -> object:
            assert name == "Excel.Application"
            return object()

        def DispatchEx(self, _name: str) -> object:
            raise AssertionError("DispatchEx must not be called")

        def Dispatch(self, _name: str) -> object:
            raise AssertionError("Dispatch must not be called")

    app = _acquire_running_excel(Client(), new_instance=False)
    assert app is not None
    with pytest.raises(ExcelComError, match="請先手動開 Excel"):
        _acquire_running_excel(Client(), new_instance=True)


def test_excel_com_source_has_no_dispatch_call():
    src = Path("lme_daily/excel_com.py").read_text(encoding="utf-8")
    assert "GetActiveObject(" in src
    assert "win32com_client.Dispatch" not in src
    assert ".DispatchEx(" not in src
    assert ".Dispatch(" not in src
