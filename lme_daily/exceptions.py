"""流程專用例外：任一步失敗都應拋出並停止，不可靜默跳過。"""


class LMEAutomationError(RuntimeError):
    """LME 每日報價流程失敗。"""


class ConfigError(LMEAutomationError):
    """config.yaml 缺失欄位、路徑無效、或檔案不存在。"""


class DateCalcError(LMEAutomationError):
    """交易日 / 3M date 計算失敗。"""


class ExcelComError(LMEAutomationError):
    """Excel COM 操作失敗（開啟、巨集、刷新、讀取）。"""


class MacroOutputError(LMEAutomationError):
    """VBA 執行後未在時限內產生預期的 yyyymmdd.xlsx。"""


class ReportBuildError(LMEAutomationError):
    """最終報告工作簿產生失敗。"""


class BbgWorkbookNotOpenError(LMEAutomationError):
    """LME BBG WORKBOOK.xlsx 沒有在 Excel 裡開著；腳本不會代為開啟。"""
