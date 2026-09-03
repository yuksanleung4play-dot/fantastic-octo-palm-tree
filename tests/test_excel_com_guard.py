"""Windows-only COM 在非 Windows 必須明確失敗。"""

from __future__ import annotations

import sys

import pytest

from lme_daily.excel_com import require_windows_excel
from lme_daily.exceptions import ExcelComError


@pytest.mark.skipif(sys.platform == "win32", reason="此測試驗證非 Windows 會被拒絕")
def test_require_windows_excel_fails_off_windows():
    with pytest.raises(ExcelComError, match="必須在已安裝 Excel"):
        require_windows_excel()
