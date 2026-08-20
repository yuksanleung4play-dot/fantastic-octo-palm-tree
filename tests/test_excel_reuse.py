"""Excel 沿用既有進程：不 DispatchEx、不誤關已斷線工作簿。"""

from __future__ import annotations

from contextlib import contextmanager
from datetime import date
from pathlib import Path
from unittest.mock import MagicMock

import pytest
import yaml

from lme_daily.excel_com import (
    RPC_E_DISCONNECTED,
    close_workbook,
    close_workbook_if_opened,
    excel_app,
    get_workbook_open_count,
    is_rpc_disconnected,
    open_workbook,
    reset_workbook_open_count,
)
from lme_daily.exceptions import ExcelComError
from lme_daily.vba_runner import run_reference_macro


class FakeComError(Exception):
    """Mimic pywintypes.com_error for RPC_E_DISCONNECTED."""

    def __init__(self, hresult: int = RPC_E_DISCONNECTED) -> None:
        self.hresult = hresult
        super().__init__(hresult, "The object invoked has disconnected from its clients.")


def _write_config(tmp_path: Path) -> Path:
    work = tmp_path / "work"
    work.mkdir()
    payload = {
        "paths": {
            "working_dir": str(work),
            "ref_workbook_name": "ref.xlsm",
            "bbg_workbook_name": "bbg.xlsx",
            "output_prefix": "LME每日報價",
        },
        "vba": {
            "macro_name": "RunDailyLME",
            "use_param_injection": True,
            "auto_closes_workbook": True,
        },
        "bloomberg": {
            "copy_range": "B3:I10",
            "bbg_sheet_name": "Promt date",
            "refresh_wait_seconds": 15,
            "calculation_timeout_seconds": 1,
        },
        "chart": {"forward_months": 27, "engine": "matplotlib"},
        "logging": {"level": "INFO", "file": ""},
    }
    path = tmp_path / "config.yaml"
    path.write_text(yaml.safe_dump(payload, allow_unicode=True), encoding="utf-8")
    return path


class _FakePythoncom:
    def CoInitialize(self) -> None:
        return None

    def CoUninitialize(self) -> None:
        return None


def test_is_rpc_disconnected():
    assert is_rpc_disconnected(FakeComError()) is True
    wrapped = ExcelComError("x")
    wrapped.__cause__ = FakeComError()
    assert is_rpc_disconnected(wrapped) is True
    assert is_rpc_disconnected(RuntimeError("nope")) is False


def test_close_workbook_swallows_rpc_disconnected():
    class Wb:
        @property
        def Name(self) -> str:
            raise FakeComError()

        def Close(self, SaveChanges: bool = False) -> None:
            raise FakeComError()

    close_workbook(Wb())
    close_workbook_if_opened(Wb(), True)


def test_excel_app_requires_running_excel(monkeypatch: pytest.MonkeyPatch):
    class Client:
        def GetActiveObject(self, _name: str) -> object:
            raise RuntimeError("MK_E_UNAVAILABLE")

        def DispatchEx(self, _name: str) -> object:
            raise AssertionError("DispatchEx must not be called")

        def Dispatch(self, _name: str) -> object:
            raise AssertionError("Dispatch must not be called")

    monkeypatch.setattr(
        "lme_daily.excel_com.import_win32com",
        lambda: (Client(), _FakePythoncom()),
    )
    with pytest.raises(ExcelComError, match="請先手動開 Excel"):
        with excel_app():
            raise AssertionError("must not enter with-block")


def test_excel_app_does_not_wrap_caller_errors(monkeypatch: pytest.MonkeyPatch):
    app = MagicMock()

    class Client:
        def GetActiveObject(self, _name: str) -> object:
            return app

        def DispatchEx(self, _name: str) -> object:
            raise AssertionError("DispatchEx must not be called")

    monkeypatch.setattr(
        "lme_daily.excel_com.import_win32com",
        lambda: (Client(), _FakePythoncom()),
    )
    with pytest.raises(RuntimeError, match="macro boom"):
        with excel_app():
            raise RuntimeError("macro boom")


def test_excel_app_rejects_new_instance(monkeypatch: pytest.MonkeyPatch):
    class Client:
        def GetActiveObject(self, _name: str) -> object:
            raise AssertionError("should fail before GetActiveObject")

        def DispatchEx(self, _name: str) -> object:
            raise AssertionError("DispatchEx must not be called")

    monkeypatch.setattr(
        "lme_daily.excel_com.import_win32com",
        lambda: (Client(), _FakePythoncom()),
    )
    with pytest.raises(ExcelComError, match="請先手動開 Excel"):
        with excel_app(new_instance=True):
            pass


def test_open_workbook_counts_only_real_open():
    reset_workbook_open_count()

    class Workbooks:
        def __iter__(self):
            return iter(())

        def Open(self, *_args: object, **_kwargs: object) -> object:
            return object()

    app = MagicMock()
    app.Workbooks = Workbooks()
    missing = Path("/tmp/definitely-missing-lme-workbook-xyz.xlsm")
    with pytest.raises(ExcelComError, match="工作簿不存在"):
        open_workbook(app, missing)
    assert get_workbook_open_count() == 0

    existing = Path(__file__).resolve()
    wb, opened = open_workbook(app, existing)
    assert opened is True
    assert wb is not None
    assert get_workbook_open_count() == 1

    class OpenBook:
        FullName = str(existing)
        Name = existing.name

    class Occupied:
        def __iter__(self):
            return iter((OpenBook(),))

        def Open(self, *_args: object, **_kwargs: object) -> object:
            raise AssertionError("Open must not be called for already-open workbook")

    app.Workbooks = Occupied()
    reused, opened_again = open_workbook(app, existing)
    assert opened_again is False
    assert reused is not None
    assert get_workbook_open_count() == 1


def test_run_reference_macro_survives_disconnected_workbook(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    from lme_daily.config import load_config

    cfg = _write_config(tmp_path)
    config = load_config(cfg)
    as_of = date(2026, 8, 20)
    dest = config.step2_workbook_path(as_of)
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(b"curve")

    class Wb:
        def __init__(self) -> None:
            self.dead = False

        @property
        def Name(self) -> str:
            if self.dead:
                raise FakeComError()
            return "ref.xlsm"

        def Activate(self) -> None:
            if self.dead:
                raise FakeComError()

        def Close(self, SaveChanges: bool = False) -> None:
            raise AssertionError("Python must not Close the reference workbook")

    wb = Wb()

    def _execute(*_args: object, **_kwargs: object) -> None:
        wb.dead = True

    @contextmanager
    def fake_excel_app(**_kwargs: object):
        yield MagicMock()

    monkeypatch.setattr("lme_daily.vba_runner.excel_app", fake_excel_app)
    monkeypatch.setattr("lme_daily.vba_runner.open_workbook", lambda *_a, **_k: (wb, True))
    monkeypatch.setattr("lme_daily.vba_runner._execute_macro", _execute)
    monkeypatch.setattr("lme_daily.vba_runner.wait_for_any_file", lambda *_a, **_k: dest)

    ready = run_reference_macro(
        config,
        as_of=as_of,
        prev_date="2026/08/19",
        three_m_date="2026/11/20",
    )
    assert ready == dest
    with pytest.raises(FakeComError):
        _ = wb.Name
