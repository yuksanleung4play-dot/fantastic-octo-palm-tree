# LME 每日報價自動化

Windows 本機流程：開啟「早班 LME reference」執行 VBA → 刷新 Bloomberg 工作簿 → 產出 `LME每日報價yyyymmdd.xlsx`（BBG 快照、六條遠期曲線、原始數據）。

> 完整流程需要 **Windows + Excel + Bloomberg Terminal**。日期計算與報告產生可在任何平台跑單元測試。

## 一鍵產生 `LME每日報價yyyymmdd.xlsx`

在檔案總管**雙擊**（不要在 cmd 裡手動打一堆殘缺指令）：

1. 第一次：雙擊 `install_deps.bat`（安裝套件，需連網）
2. 之後每天：雙擊 **`RUN_LME.bat`**（或 `RUN_LME.vbs`）

成功後會印出 `OUTPUT=` 下一行就是 `LME每日報價yyyymmdd.xlsx` 的完整路徑，並用 Excel 打開。

請先改好 `config.yaml` 的 `paths.working_dir`，並登入 Bloomberg、開得了 Excel。

**不要**把 `pywinauto`、`requirements.txt` 當成指令執行。完整說明見 `HOW_TO_RUN.txt`。

舊版 bat 若出現 `嘿濃 echo`：那是 UTF-8 BOM 被繁中 cmd 誤讀，請改用現在這個無 BOM 的 `RUN_LME.bat`。資料夾路徑若含 `&`（例如 `LME --Form & Sheet`），務必雙擊 bat，不要手動 `cd` 未加引號的路徑。

## 快速開始

```text
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
pip install pywin32 pywinauto          # Windows 才需要
```

編輯 `config.yaml` 與 `holidays.yaml`（路徑、巨集名稱、公休日），確認 `working_dir` 裡已有：

- `早班_LME_reference_2024.xlsm`（或你在 config 填的檔名）
- `LME BBG WORKBOOK.xlsx`

執行：

```text
python -m lme_daily --config config.yaml
```

常用參數：

| 參數 | 說明 |
|------|------|
| `--as-of YYYY-MM-DD` | 覆寫執行日（影響上日 / 3M / 輸出檔名） |
| `--dry-run` | 只檢查設定並計算日期，不開 Excel |
| `--skip-vba` | 使用 `working_dir` 裡既有的 `yyyymmdd.xlsx` |
| `--skip-bbg` | 跳過 Bloomberg 刷新 |

日誌會打到 console，並寫入 `working_dir` 下的 `lme_daily.log`。

## 目錄

```text
config.yaml
holidays.yaml
HOW_TO_RUN.txt
RUN_LME.bat                # 請雙擊這個（無 BOM，純英文）
RUN_LME.vbs
install_deps.bat           # 第一次安裝套件
Generate_LME_Daily.bat     # 與 RUN_LME.bat 相同
generate_lme_daily.py
run_daily.py
lme_daily/
  main.py              # 流程串接
  config.py            # 讀 YAML、路徑檢查
  dates.py             # 上日日期 / 3M date
  excel_com.py         # win32com Excel 封裝
  vba_runner.py        # 參數注入 或 pywinauto InputBox
  bbg_fetch.py         # RefreshAll + Value2
  report_builder.py    # 三個 sheet；matplotlib 或 xlsxwriter
examples/RunDailyLME_param_wrapper.bas
tests/
```

## 日期規則（`dates.calc_lme_dates`）

- **上日日期**：從執行日前一天往前，找最近一個 Settlement Business Day（排除週末 + `holiday_list`）。
- **3M date**：執行日 + 3 個自然月（同一天，月底由 `relativedelta` 夾住）。若不是交易日則先順延；若順延跨月則改往前推到上一個交易日。

公休日請填 `holidays.yaml` 或 `config.yaml` 的 `holidays.dates`。未填時**只排除週末**。

## VBA 兩種模式

`vba.use_param_injection: true`（預設）會呼叫：

```text
Application.Run("RunDailyLME", 上日日期, 3M date)
```

巨集必須接受兩個 Optional 參數。可參考 `examples/RunDailyLME_param_wrapper.bas`。

若巨集只能走 InputBox，改成 `use_param_injection: false`，改由 `pywinauto` 自動填兩個彈窗並按 Enter。

巨集完成後必須在 `working_dir` 產生 `yyyymmdd.xlsx`（例如 `20260819.xlsx`），腳本會 polling 等到檔案就緒。

## 報告內容

同一個工作簿三個 sheet：

1. **BBG快照** — Bloomberg 工作表 `copy_range`（預設 `B3:I10`）的 `.Value2`
2. **遠期走勢圖** — 以 cash date 起 `chart.forward_months`（預設 27）個月為視窗，六個品種各一張圖。`NI` / `SN` 空值會 `dropna`，**不會填 0**
3. **原始數據** — 步驟二檔案的完整內容（含 27 個月以後）

`chart.engine`：

- `matplotlib`：PNG 嵌入，2 欄 × 3 列
- `xlsxwriter`：Excel 原生可互動折線圖（資料寫在隱藏 sheet `_chart_data`）

## 尚待本機確認

- `vba.macro_name` 真實名稱，以及巨集是否支援參數覆蓋 InputBox
- `bloomberg.bbg_sheet_name`（預設 `Promt date`）
- LME 公休日清單（`holidays.yaml`）

## 測試

```text
pip install pytest
pytest
```

COM / Bloomberg 步驟在非 Windows 環境會明確報錯並停止，不會假裝成功。
