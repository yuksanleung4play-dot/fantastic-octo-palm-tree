# fantastic-octo-palm-tree

## Invoice 整理工具 (`organize_invoices.py`)

掃描資料夾內所有 invoice PDF，逐張抓取下列欄位並彙整成一份 CSV：

| 欄位 | 說明 | 範例 |
| --- | --- | --- |
| `status` | 交易狀態；PDF 含 `***** PLEASE VOID THIS CONFIRMATION, TRADE IS NOT VALID *****` 時為 `void`，否則 `valid` | `void` |
| `trade_id` | Trade ID | `3394221 S` |
| `symbol` | 合約代號 (Symbo) | `SMT` |
| `qty` | 數量 (Qty / Barrel) | `31000` |
| `price` | 價格 (Price) | `0` |
| `commission` | 佣金 | `186.00` |
| `commission_rate` | 佣金費率 (per unit) | `0.006` |

支援單張 invoice 內多個 leg（每個 leg 各輸出一列）。

### 安裝

```bash
pip install -r requirements.txt
```

### 使用

```bash
# 整理目前目錄
python organize_invoices.py

# 指定資料夾
python organize_invoices.py ./invoices

# 指定輸出檔名
python organize_invoices.py ./invoices -o summary.csv
```

執行後會在資料夾內產生 `invoice_summary.csv`（UTF-8 BOM，可直接用 Excel 開啟），並在終端機印出彙整表格。

範例 invoice 放在 `sample_invoices/`，可用來測試：

```bash
python organize_invoices.py sample_invoices
```
