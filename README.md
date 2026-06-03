# 週報生成器 Weekly Report Generator

根據**郵箱發件箱**與**額外輸入的資料**，自動整理成一份結構化的週報。

週報固定包含三大區塊：

1. **姓名、日期、部門**
2. **日常工作內容**（由發件箱郵件自動彙整，並可手動補充）
3. **其他事項**（下週計劃、需協調事項、請假等）

支援輸出 **純文字 / Markdown / HTML** 三種格式。

---

## 運作原理

```
                ┌──────────────────────┐
  發件箱  ─────▶│  email_source        │  IMAP 或 本地 .mbox/.eml
                │  讀取已寄出郵件        │
                └──────────┬───────────┘
                           │
  額外輸入 ───────────────▶│  report_builder  依日期分組、合併同主題、
 (姓名/部門/其他事項/手動)   │                   去除 Re:/Fwd: 前綴
                           ▼
                ┌──────────────────────┐
                │  renderers           │  text / markdown / html
                └──────────────────────┘
```

- 郵件依「寄出日期」分組，做為當天的日常工作內容。
- 同一天裡同主題（去掉 `Re:`／`Fwd:` 後相同）的郵件會合併成一條，避免同一串對話重複佔行。
- 「其他事項」與「手動工作項目」由設定檔或命令列提供。

---

## 安裝

核心功能只用 Python 標準函式庫（`imaplib`、`email`、`mailbox`），無需安裝任何套件。

若要使用 **YAML 設定檔**，請安裝 PyYAML（也可改用 JSON 設定檔，免安裝）：

```bash
pip install -r requirements.txt
```

需求：Python 3.8+。

---

## 快速開始

### 1. 用本地匯出的發件箱檔（最簡單，先試用範例）

```bash
python -m weekly_report \
  --mbox samples/sent.mbox \
  --name 王小明 --department 研發部 \
  --start 2026-06-01 --end 2026-06-07 \
  --other "下週準備 Q2 簡報" --other "6/12 請假一天" \
  -f text
```

### 2. 用設定檔（推薦，可放 IMAP 帳密、其他事項、手動項目）

```bash
cp config.example.yaml config.yaml   # 修改成自己的資料
python -m weekly_report --config config.yaml -f html -o report.html
```

### 3. 直接連 IMAP 讀「寄件備份」

```bash
export WEEKLY_REPORT_IMAP_PASSWORD='你的密碼或App專用密碼'
python -m weekly_report \
  --imap-host imap.example.com --imap-user you@example.com \
  --name 王小明 --department 研發部 -f markdown
```

> 密碼建議透過環境變數 `WEEKLY_REPORT_IMAP_PASSWORD` 或設定檔的 `password_env` 提供，
> 不要寫死在命令列。Gmail 等服務需使用「應用程式專用密碼」。

### 4. 不接郵箱，純手動 / 互動輸入

```bash
python -m weekly_report --name 王小明 --department 研發部 --interactive
```

---

## 郵件來源

| 來源 | 說明 | 參數 |
| --- | --- | --- |
| `.mbox` | 多數郵件軟體可匯出的標準格式 | `--mbox path.mbox` |
| `.eml` | 單封郵件，或內含多封 `.eml` 的資料夾 | `--eml path` |
| IMAP | 線上讀取「寄件備份」資料夾 | `--imap-host/--imap-user/...` |

IMAP 的寄件備份資料夾名稱會**自動偵測**常見命名（Sent、寄件備份、已發送、`[Gmail]/Sent Mail`…）；
若偵測失敗可用 `--sent-folder` 手動指定。

---

## 常用參數

| 參數 | 說明 |
| --- | --- |
| `--config` | 設定檔（`.yaml` 或 `.json`） |
| `--name` / `--department` / `--title` | 基本資訊（可覆蓋設定檔） |
| `--start` / `--end` | 指定起訖日期（`YYYY-MM-DD`） |
| `--last-week` | 改用上一週（預設為本週，週一至週日） |
| `--other 事項` | 其他事項，可重複多次 |
| `--interactive` | 互動模式，逐項詢問缺少的資訊 |
| `--include-empty-days` | 保留沒有工作內容的日期 |
| `-f, --format` | 輸出格式：`text` / `markdown` / `html` |
| `-o, --output` | 輸出到檔案（預設印到畫面） |

完整說明：`python -m weekly_report --help`

---

## 設定檔格式

見 [`config.example.yaml`](config.example.yaml)。重點欄位：

```yaml
name: 王小明
department: 研發部
period_start: "2026-06-01"   # 或留空改用 week_offset
period_end: "2026-06-07"

email:
  source: mbox
  path: samples/sent.mbox
  # 或 IMAP：source: imap / host / username / password_env / sent_folder

other_matters:
  - 下週一部門例會需準備簡報。

manual_items:               # 自動補充某天郵件未涵蓋的工作
  "2026-06-03":
    - 完成新版報表模組的需求評審。
```

---

## 專案結構

```
weekly_report/
├── __init__.py
├── __main__.py        # 讓 `python -m weekly_report` 可執行
├── models.py          # 資料模型（dataclass）
├── email_source.py    # 發件箱讀取：IMAP + 本地 mbox/eml
├── report_builder.py  # 組裝週報（分組、去重、合併手動項目）
├── renderers.py       # 輸出：text / markdown / html
├── config.py          # 設定檔載入
└── cli.py             # 命令列介面
samples/sent.mbox      # 範例發件箱
tests/                 # 單元測試
config.example.yaml    # 設定檔範例
```

---

## 測試

```bash
python -m unittest discover -s tests -v
```

---

## 安全提醒

- 請勿把含有真實帳密的 `config.yaml` 提交到版本庫（已列入 `.gitignore`）。
- 優先使用環境變數或應用程式專用密碼。
