#!/usr/bin/env python3
"""整理資料夾裡的 invoice (PDF)，抓取 Trade ID、Symbol、Qty、Price、Commission。

範本：Oil Brokerage Limited / Shanxi Securities 之 Trade Confirmation。

用法：
    python organize_invoices.py [資料夾路徑] [-o 輸出.csv]

範例：
    python organize_invoices.py ./invoices
    python organize_invoices.py ./invoices -o summary.csv
"""
from __future__ import annotations

import argparse
import csv
import re
import sys
from pathlib import Path

try:
    import pdfplumber
except ImportError:  # pragma: no cover
    sys.exit("缺少套件 pdfplumber，請先執行：pip install -r requirements.txt")


# --- 各欄位的擷取規則 ------------------------------------------------------

# 作廢標記，例： "***** PLEASE VOID THIS CONFIRMATION, TRADE IS NOT VALID *****"
RE_VOID = re.compile(r"PLEASE VOID THIS CONFIRMATION", re.IGNORECASE)

# Date: 05/11/2026
RE_DATE = re.compile(r"Date:\s*(\d{1,2}/\d{1,2}/\d{2,4})")

# Trade ID: 3394221 S  (其後接 Version 號，需排除)
RE_TRADE_ID = re.compile(r"Trade ID:\s*([0-9]+(?:\s+[A-Z]+)?)")

# Leg 資料列，倒數兩個數字為 qty、price。leg 編號後須接空白與非數字，
# 以排除 "06/01/2026 - 06/30/2026 ..." 之類的日期/結算列。
#   例1： "1 SMT ICE (Euro) International Futures 31,000 0"
#   例2： "1 Unleaded (Platts) Mini SMV ICE (Euro) International Futures 5,300 0"
RE_LEG = re.compile(r"^\s*(\d+)\s+\S.*?([\d,]+)\s+(\d+(?:\.\d+)?)\s*$")

# leg 列開頭的 symbol（全大寫代號），例： "1 SMT ..."
RE_LEG_SYMBOL = re.compile(r"^\s*\d+\s+([A-Z]{2,6})\b")

# 文件中括號內的全大寫合約代號，例： "(SMT)"、"(SMV)"
RE_PAREN_SYMBOL = re.compile(r"\(([A-Z]{2,6})\)")

# 每個 leg 的佣金，例： "Commission: 186.00 (Leg 1S) = USD 186.00"
RE_COMMISSION_LEG = re.compile(r"Commission:\s*([\d,]+(?:\.\d+)?)\s*\(Leg\s*(\d+)")

# 總佣金 (fallback)
RE_COMMISSION = re.compile(r"Commission:\s*([\d,]+(?:\.\d+)?)")

# 佣金費率，例： "Commission Rate per Barrel USD 0.006"
RE_COMMISSION_RATE = re.compile(r"Commission Rate per\s+\w+\s+[A-Z]{3}\s+([\d.]+)")

# 後備：symbol 取自 "Future (SMT)"，qty 取自 "Qty/Period: 31,000 Barrel"
RE_SYMBOL_FALLBACK = re.compile(r"Future\s*\(([A-Z]{2,6})\)")
RE_QTY_FALLBACK = re.compile(r"Qty/Period:\s*([\d,]+)")


def extract_text(pdf_path: Path) -> str:
    """讀出 PDF 全文 (多頁合併)。"""
    parts: list[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            parts.append(page.extract_text() or "")
    return "\n".join(parts)


def _num(value: str) -> str:
    """去掉千分位逗號。"""
    return value.replace(",", "").strip()


def parse_invoice(text: str, source: str) -> list[dict]:
    """從單張 invoice 全文解析出每個 leg 的欄位，回傳 list[dict]。"""
    status = "void" if RE_VOID.search(text) else "valid"

    date = ""
    if m := RE_DATE.search(text):
        date = m.group(1).strip()

    trade_id = ""
    if m := RE_TRADE_ID.search(text):
        trade_id = m.group(1).strip()

    commission_rate = ""
    if m := RE_COMMISSION_RATE.search(text):
        commission_rate = m.group(1)

    leg_commissions = {
        leg: _num(amount) for amount, leg in RE_COMMISSION_LEG.findall(text)
    }
    total_commission = ""
    if m := RE_COMMISSION.search(text):
        total_commission = _num(m.group(1))

    # 文件層級的合約代號 (依出現順序、去重)，作為 leg 列抓不到 symbol 時的後備
    doc_symbols = list(dict.fromkeys(RE_PAREN_SYMBOL.findall(text)))

    legs: list[dict] = []
    for line in text.splitlines():
        if lm := RE_LEG.match(line):
            leg_no, qty, price = lm.groups()
            idx = len(legs)
            # symbol 優先序：leg 列開頭代號 -> 同列括號代號 -> 文件層級代號
            if sm := RE_LEG_SYMBOL.match(line):
                symbol = sm.group(1)
            elif sp := RE_PAREN_SYMBOL.search(line):
                symbol = sp.group(1)
            elif idx < len(doc_symbols):
                symbol = doc_symbols[idx]
            else:
                symbol = doc_symbols[0] if doc_symbols else ""
            legs.append(
                {
                    "leg": leg_no,
                    "symbol": symbol,
                    "qty": _num(qty),
                    "price": price.strip(),
                }
            )

    # 後備：完全沒抓到 leg 資料列時，改用標籤式欄位
    if not legs:
        symbol = m.group(1) if (m := RE_SYMBOL_FALLBACK.search(text)) else ""
        qty = _num(m.group(1)) if (m := RE_QTY_FALLBACK.search(text)) else ""
        if symbol or qty:
            legs.append({"leg": "1", "symbol": symbol, "qty": qty, "price": ""})

    rows: list[dict] = []
    for leg in legs:
        commission = leg_commissions.get(leg["leg"], "")
        if not commission and len(legs) == 1:
            commission = total_commission
        rows.append(
            {
                "file": source,
                "status": status,
                "date": date,
                "trade_id": trade_id,
                "symbol": leg["symbol"],
                "qty": leg["qty"],
                "price": leg["price"],
                "commission": commission,
                "commission_rate": commission_rate,
            }
        )

    # 連 leg 都沒有 (格式不符)，仍輸出一列以利人工檢查
    if not rows:
        rows.append(
            {
                "file": source,
                "status": status,
                "date": date,
                "trade_id": trade_id,
                "symbol": "",
                "qty": "",
                "price": "",
                "commission": total_commission,
                "commission_rate": commission_rate,
            }
        )
    return rows


FIELDS = [
    "file",
    "status",
    "date",
    "trade_id",
    "symbol",
    "qty",
    "price",
    "commission",
    "commission_rate",
]


def organize(folder: Path, output: Path) -> list[dict]:
    pdf_files = sorted(folder.glob("*.pdf")) + sorted(folder.glob("*.PDF"))
    pdf_files = sorted(set(pdf_files))
    if not pdf_files:
        print(f"在 {folder} 找不到任何 PDF 檔。", file=sys.stderr)
        return []

    all_rows: list[dict] = []
    for pdf_path in pdf_files:
        try:
            text = extract_text(pdf_path)
            rows = parse_invoice(text, pdf_path.name)
            all_rows.extend(rows)
            print(f"[OK]   {pdf_path.name} -> {len(rows)} 筆")
        except Exception as exc:  # noqa: BLE001
            print(f"[FAIL] {pdf_path.name}: {exc}", file=sys.stderr)

    if all_rows:
        with output.open("w", newline="", encoding="utf-8-sig") as fh:
            writer = csv.DictWriter(fh, fieldnames=FIELDS)
            writer.writeheader()
            writer.writerows(all_rows)
        print(f"\n已彙整 {len(all_rows)} 筆，輸出 -> {output}")

    return all_rows


def print_table(rows: list[dict]) -> None:
    if not rows:
        return
    widths = {f: max(len(f), *(len(str(r.get(f, ""))) for r in rows)) for f in FIELDS}
    line = " | ".join(f.ljust(widths[f]) for f in FIELDS)
    print("\n" + line)
    print("-" * len(line))
    for r in rows:
        print(" | ".join(str(r.get(f, "")).ljust(widths[f]) for f in FIELDS))


def main() -> None:
    parser = argparse.ArgumentParser(description="整理資料夾裡的 invoice PDF，抓取交易欄位。")
    parser.add_argument("folder", nargs="?", default=".", help="invoice 所在資料夾 (預設目前目錄)")
    parser.add_argument(
        "-o", "--output", default="invoice_summary.csv", help="輸出 CSV 檔名"
    )
    args = parser.parse_args()

    folder = Path(args.folder).expanduser().resolve()
    if not folder.is_dir():
        sys.exit(f"資料夾不存在：{folder}")

    output = Path(args.output)
    if not output.is_absolute():
        output = folder / output

    rows = organize(folder, output)
    print_table(rows)


if __name__ == "__main__":
    main()
