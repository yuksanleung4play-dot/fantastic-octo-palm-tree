# LME 每日報價自動化

Windows 本機流程：在已開啟的 Excel 執行早班 LME VBA → 讀取已手動打開的 Bloomberg 工作簿 → 產出 `LME每日報價yyyymmdd.xlsx`（BBG 快照、六條遠期曲線、原始數據、合併版面）與 PDF。

> 完整流程需要 **Windows + Excel + Bloomberg Terminal**。日期計算與報告產生可在任何平台跑單元測試。

## 一鍵產生 `LME每日報價yyyymmdd.xlsx`

在檔案總管**雙擊**（不要在 cmd 裡手動打一堆殘缺指令）：

1. 第一次：雙擊 `install_deps.bat`（安裝套件，需連網）
2. 之後每天：雙擊 **`RUN_LME.bat`**（或 `RUN_LME.vbs`）

成功後會印出 `OUTPUT=` 下一行就是 `LME每日報價yyyymmdd.xlsx` 的完整路徑，並用 Excel 打開。

請確認巨集會彈出兩個 InputBox（上日、3M）。`lme_main` 若沒有參數，程式會自動改填彈窗，不必改 config。

## Bloomberg Terminal 被鎖

**不能、也不該自動解鎖**（Bloomberg 安全機制與使用條款）。腳本以前用 `DispatchEx` 另開 Excel 再 `Quit()`，最容易把 Terminal 鎖上。

現在預設：

1. 先手動登入並解鎖 Bloomberg Terminal
2. **先開 Excel**（讓 Bloomberg 外掛載入在這個 Excel）
3. 腳本會自動開啟 `LME BBG WORKBOOK.xlsx`（若尚未開啟）；開啟後等待 `bloomberg.refresh_wait_seconds`（預設 15 秒）再讀取，讀取後不會關閉
4. 再跑 `RUN_LME.bat` — 沿用該 Excel，**結束不 Quit**、**不另開新 Excel 進程**

`config.yaml` 的 `bloomberg.source`：

| 值 | 做法 |
|----|------|
| `excel` | 已開就沿用、未開就 Open；等待後 RefreshAll；讀完不 Close（預設） |
| `cached` | 只讀目前儲存格，不 Refresh（最不易鎖） |
| `blpapi` | Python Desktop API，不經 Excel 外掛（`pip install blpapi`；Terminal 仍須已登入未鎖定） |

真正無人值守、不開 Terminal，只能改用 Bloomberg **Data License / B-PIPE**（需另外簽約）。

**路徑請用單引號**，否則 YAML 會把 `\Dealing` 當成非法跳脫（`unknown escape character 'D'`）：

```yaml
working_dir: '\\192.168.89.167\Dealing\Dealing Department - New\...'
```

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
| `--skip-vba` | 使用 `working_dir\yyyymmdd\` 裡既有的 `yyyymmdd.xlsx`（若只在根目錄也會搬入） |
| `--skip-bbg` | 跳過 Bloomberg 刷新 |

日誌會打到 console；`logging.file` 若有填會寫一份到 `working_dir`，另外**每天一份** `lme_daily.log` 一定寫進最終報告資料夾 `run_dir`。

## 兩個資料夾：`vba_dir` vs `run_dir`

整個流程有兩個用途不同的資料夾，不要混用：

| 變數 | 位置 | 放什麼 |
|------|------|--------|
| `vba_dir` | 永遠 `working_dir\yyyymmdd\` | VBA 巨集產生的中繼檔 `yyyymmdd.xlsx`（例如 `20260820.xlsx`） |
| `run_dir` | `output_dir` 留空 → 與 `vba_dir` 相同；有填 → `output_dir\yyyymmdd\` | `LME每日報價yyyymmdd.xlsx`、對應 `.pdf`、當日 `lme_daily.log` |

```yaml
paths:
  working_dir: '\\192.168.89.167\Dealing\...'
  output_prefix: "LME每日報價"
  output_dir: ''    # 留空 = 最終報告也放 working_dir\yyyymmdd\
                    # 填了路徑 = 最終報告改放 output_dir\yyyymmdd\
```

兩個資料夾都會在流程一開始 `mkdir(parents=True, exist_ok=True)`（同一天重跑不會報錯）。來源工作簿（`早班_LME_reference_2024.xlsm`、`LME BBG WORKBOOK.xlsx`）永遠讀自 `working_dir` 根目錄。

**關鍵：** 不管 `output_dir` 有沒有填，「遠期走勢圖」與「原始數據」的資料來源一律是 `working_dir\yyyymmdd\yyyymmdd.xlsx`。原始數據 sheet 用**絕對路徑** Excel 外部連結，避免報告跟中繼檔不在同一資料夾時斷鏈。若該中繼檔不存在，流程會中斷，不會產出缺資料的報告。

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
  report_builder.py    # 四個 sheet（含合併版面）；matplotlib 或 xlsxwriter
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

巨集完成後必須產生 `yyyymmdd.xlsx`。程式會先在 `working_dir\yyyymmdd\`（`vba_dir`）等檔；若巨集仍寫到 `working_dir` 根目錄，會自動搬進 `vba_dir`。最終報告**不會**把這份中繼檔跟著 `output_dir` 搬走。

## 報告內容

同一個工作簿四個可見 sheet：

1. **BBG快照** — 腳本會自動開啟 `LME BBG WORKBOOK.xlsx`（若尚未開啟），開啟後等待設定秒數再讀 `copy_range`（預設 `B3:I10`）的 `.Value2`，讀取後不會關閉。字串若以 `N/A` 開頭（例如 `N/A Field Not Applicable`）會正規化成純 `N/A`；空白 / `None` / 數字維持原樣，不會填字。
2. **遠期走勢圖** — 以 cash date 起 `chart.forward_months`（預設 27）個月為視窗，六個品種各一張圖。資料來自 `vba_dir\yyyymmdd.xlsx`。`NI` / `SN` 空值會 `dropna`，**不會填 0**
3. **原始數據** — 以絕對路徑外部連結指向 `vba_dir\yyyymmdd.xlsx` 的完整內容（含 27 個月以後）
4. **合併版面** — 把上述三塊排在同一橫向 A3 版面（深藍/白底、中英雙語標題、頁碼「第 n 頁，共 N 頁」）。Windows 下用 `ExportAsFixedFormat` 匯出 `LME每日報價yyyymmdd.pdf`

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
