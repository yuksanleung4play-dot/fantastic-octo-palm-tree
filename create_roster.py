#!/usr/bin/env python3
"""Create a cross-year reusable department shift roster Excel template."""

from datetime import date, timedelta

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    PatternFill,
    Side,
    NamedStyle,
    Protection,
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
DISPLAY_YEAR = 2026
BASE_DATE = date(2026, 6, 29)  # Permanent rotation anchor (Monday)
AL_QUOTA = 12
NUM_WEEKS = 53

EMPLOYEES = [
    # name, role (S=fixed, R=rotate, V=vacant), offset
    ("ARTHUR", "S", None),
    ("PATRICK", "R", 0),
    ("CARREY", "R", 4),
    ("CESC", "R", 8),
    ("LUCAS", "R", 12),
    ("", "V", 0),  # vacant slot ready for future hire
]

HOLIDAYS_2026 = [
    (date(2026, 1, 1), "元旦"),
    (date(2026, 2, 17), "農曆年初一"),
    (date(2026, 2, 18), "農曆年初二"),
    (date(2026, 2, 19), "農曆年初三"),
    (date(2026, 4, 3), "耶穌受難節前一日"),
    (date(2026, 4, 4), "耶穌受難節"),
    (date(2026, 4, 6), "復活節星期一"),
    (date(2026, 5, 1), "勞動節"),
    (date(2026, 5, 25), "佛誕"),
    (date(2026, 6, 19), "端午節"),
    (date(2026, 7, 1), "香港特別行政區成立紀念日"),
    (date(2026, 9, 26), "中秋節翌日"),
    (date(2026, 10, 1), "國慶日"),
    (date(2026, 10, 19), "重陽節"),
    (date(2026, 12, 25), "聖誕節"),
    (date(2026, 12, 26), "聖誕節後第一個周日"),
]

# Colors
FILL_HEADER = PatternFill("solid", fgColor="1F4E79")
FILL_SUBHDR = PatternFill("solid", fgColor="2E75B6")
FILL_SETTINGS = PatternFill("solid", fgColor="D6EAF8")
FILL_QUICK = PatternFill("solid", fgColor="FFF2CC")
FILL_QUICK_HDR = PatternFill("solid", fgColor="F4B183")
FILL_S = PatternFill("solid", fgColor="FFFFFF")
FILL_A = PatternFill("solid", fgColor="D9D9D9")
FILL_P = PatternFill("solid", fgColor="C6EFCE")
FILL_N = PatternFill("solid", fgColor="BDD7EE")
FILL_AL = PatternFill("solid", fgColor="FFFF00")
FILL_SL = PatternFill("solid", fgColor="F4B183")
FILL_HOLIDAY = PatternFill("solid", fgColor="E2D5F1")
FILL_VACANT = PatternFill("solid", fgColor="F2F2F2")
FILL_STAT = PatternFill("solid", fgColor="E2EFDA")
FILL_WARN = PatternFill("solid", fgColor="FF6B6B")
FILL_INPUT = PatternFill("solid", fgColor="FFF9E6")
FILL_ANCHOR = PatternFill("solid", fgColor="FCE4D6")

# Cell refs on 設定 sheet
YEAR_CELL = "'設定'!$B$3"
BASE_CELL = "'設定'!$B$4"
QUOTA_CELL = "'設定'!$B$5"
EMP_FIRST_ROW = 9  # 設定!B9:B14 names

FONT_WHITE = Font(name="Microsoft JhengHei", bold=True, color="FFFFFF", size=11)
FONT_TITLE = Font(name="Microsoft JhengHei", bold=True, size=16, color="1F4E79")
FONT_HDR = Font(name="Microsoft JhengHei", bold=True, size=10, color="FFFFFF")
FONT_NORMAL = Font(name="Microsoft JhengHei", size=10)
FONT_BOLD = Font(name="Microsoft JhengHei", bold=True, size=10)
FONT_SMALL = Font(name="Microsoft JhengHei", size=9)
FONT_NOTE = Font(name="Microsoft JhengHei", size=9, italic=True, color="C00000")

THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

FIRST_WEEK_COL = 8  # column H
STAT_COLS = {"al_used": 2, "al_left": 3, "sl_used": 4, "comp_bal": 5, "comp_open": 6, "comp_year": 7}


def week_start_formula(week_index: int) -> str:
    """First Monday on/after Jan 1 of display year, plus week_index*7; blank if past Dec 31."""
    # Matches examples like 1/5, 1/12, 1/19 for 2026
    first = (
        f"DATE({YEAR_CELL},1,1)"
        f"+MOD(8-WEEKDAY(DATE({YEAR_CELL},1,1),2),7)"
    )
    week = f"({first})+{week_index}*7"
    return f'=IF(({week})>DATE({YEAR_CELL},12,31),"",{week})'


def build_workbook():
    wb = Workbook()

    # =====================================================================
    # Sheet: 設定
    # =====================================================================
    ws_set = wb.active
    ws_set.title = "設定"

    ws_set["A1"] = "部門更表 — 設定區"
    ws_set["A1"].font = FONT_TITLE
    ws_set.merge_cells("A1:F1")

    ws_set["A3"] = "顯示年份"
    ws_set["A3"].font = FONT_BOLD
    ws_set["B3"] = DISPLAY_YEAR
    ws_set["B3"].font = Font(name="Microsoft JhengHei", bold=True, size=14, color="C00000")
    ws_set["B3"].fill = FILL_INPUT
    ws_set["B3"].border = THIN
    ws_set["B3"].alignment = CENTER
    ws_set["C3"] = "← 每年年頭只改此數字，全年週份日期與班次自動重算"
    ws_set["C3"].font = FONT_NOTE

    ws_set["A4"] = "輪班基準日（永久固定）"
    ws_set["A4"].font = FONT_BOLD
    ws_set["B4"] = BASE_DATE
    ws_set["B4"].number_format = "YYYY-MM-DD"
    ws_set["B4"].fill = FILL_ANCHOR
    ws_set["B4"].border = THIN
    ws_set["B4"].font = FONT_BOLD
    ws_set["C4"] = "← 永遠不要修改！所有輪班以此日為錨點跨年連續計算"
    ws_set["C4"].font = FONT_NOTE

    ws_set["A5"] = "年假額度（天／人／年）"
    ws_set["B5"] = AL_QUOTA
    ws_set["B5"].fill = FILL_INPUT
    ws_set["B5"].border = THIN
    ws_set["C5"] = "← 可調整；剩餘年假 = 額度 − 已用年假（按顯示年份）"
    ws_set["C5"].font = FONT_SMALL

    ws_set["A6"] = "說明"
    ws_set["B6"] = (
        "16週循環：4週A → 4週P → 4週N → 4週P。offset 錯開 0/4/8/12，確保每週 A/P/N 皆有人。"
    )
    ws_set.merge_cells("B6:F6")
    ws_set["B6"].font = FONT_SMALL

    # Employee table header (row 8); employee data rows 9–14
    headers = ["編號", "員工姓名", "班次類型", "初始Offset", "說明"]
    for i, h in enumerate(headers, 1):
        cell = ws_set.cell(8, i, h)
        cell.fill = FILL_HEADER
        cell.font = FONT_HDR
        cell.alignment = CENTER
        cell.border = THIN

    role_labels = {
        "S": "S＝固定正常班（不輪班）",
        "R": "R＝參與 A/P/N 輪班",
        "V": "V＝空缺（留空待啟用）",
    }
    notes = [
        "固定 S 班，全年不變",
        "輪班 offset=0（循環起點）",
        "輪班 offset=4（錯開一段）",
        "輪班 offset=8（錯開兩段）",
        "輪班 offset=12（錯開三段）",
        "填入姓名後，將類型改為 R 並設定 offset（建議 0/4/8/12 錯開）",
    ]

    for i, (name, role, offset) in enumerate(EMPLOYEES):
        row = EMP_FIRST_ROW + i
        ws_set.cell(row, 1, i + 1).border = THIN
        ws_set.cell(row, 1).alignment = CENTER

        name_cell = ws_set.cell(row, 2, name if name else None)
        name_cell.fill = FILL_INPUT
        name_cell.border = THIN
        name_cell.font = FONT_BOLD
        name_cell.alignment = CENTER

        role_cell = ws_set.cell(row, 3, role)
        role_cell.fill = FILL_INPUT
        role_cell.border = THIN
        role_cell.alignment = CENTER

        off_cell = ws_set.cell(row, 4, offset if offset is not None else None)
        off_cell.fill = FILL_INPUT
        off_cell.border = THIN
        off_cell.alignment = CENTER

        note_cell = ws_set.cell(row, 5, notes[i])
        note_cell.font = FONT_SMALL
        note_cell.border = THIN

        for c in range(1, 6):
            ws_set.cell(row, c).fill = FILL_SETTINGS if role != "V" else FILL_VACANT
            if c in (2, 3, 4):
                ws_set.cell(row, c).fill = FILL_INPUT

    # Named-style legend
    ws_set["A16"] = "班次代碼說明"
    ws_set["A16"].font = FONT_BOLD
    legends = [
        ("S", "正常班（ARTHUR 固定）", FILL_S),
        ("A", "早班", FILL_A),
        ("P", "中班", FILL_P),
        ("N", "晚班", FILL_N),
        ("AL", "年假 Annual Leave（手動輸入覆蓋）", FILL_AL),
        ("SL", "病假 Sick Leave（手動輸入覆蓋）", FILL_SL),
    ]
    for i, (code, desc, fill) in enumerate(legends):
        ws_set.cell(17 + i, 1, code).fill = fill
        ws_set.cell(17 + i, 1).border = THIN
        ws_set.cell(17 + i, 1).alignment = CENTER
        ws_set.cell(17 + i, 1).font = FONT_BOLD
        ws_set.cell(17 + i, 2, desc).font = FONT_NORMAL

    ws_set["A24"] = "輪班計算公式（供參考）"
    ws_set["A24"].font = FONT_BOLD
    ws_set["A25"] = (
        "週序 = INT((該週起始日 − 輪班基準日) / 7) + Offset；"
        "位置 = MOD(週序, 16)；"
        "0–3→A，4–7→P，8–11→N，12–15→P"
    )
    ws_set.merge_cells("A25:F25")
    ws_set["A25"].font = FONT_SMALL

    ws_set.column_dimensions["A"].width = 22
    ws_set.column_dimensions["B"].width = 16
    ws_set.column_dimensions["C"].width = 12
    ws_set.column_dimensions["D"].width = 12
    ws_set.column_dimensions["E"].width = 55
    ws_set.column_dimensions["F"].width = 20

    # Data validation for role
    dv_role = DataValidation(type="list", formula1='"S,R,V"', allow_blank=True)
    ws_set.add_data_validation(dv_role)
    dv_role.add(f"C{EMP_FIRST_ROW}:C{EMP_FIRST_ROW+5}")

    # =====================================================================
    # Sheet: 公眾假期
    # =====================================================================
    ws_hol = wb.create_sheet("公眾假期")
    ws_hol["A1"] = "公眾假期資料表"
    ws_hol["A1"].font = FONT_TITLE
    ws_hol.merge_cells("A1:C1")
    ws_hol["A2"] = (
        "每年年頭在此表底部追加新一年假期即可；主表標色與補假計算會自動延續。請勿刪除表頭。"
    )
    ws_hol["A2"].font = FONT_NOTE
    ws_hol.merge_cells("A2:C2")

    for i, h in enumerate(["日期", "假期名稱", "年份（自動）"], 1):
        cell = ws_hol.cell(4, i, h)
        cell.fill = FILL_HEADER
        cell.font = FONT_HDR
        cell.alignment = CENTER
        cell.border = THIN

    for i, (d, name) in enumerate(HOLIDAYS_2026):
        row = 5 + i
        ws_hol.cell(row, 1, d).number_format = "YYYY-MM-DD"
        ws_hol.cell(row, 1).border = THIN
        ws_hol.cell(row, 1).fill = FILL_HOLIDAY
        ws_hol.cell(row, 2, name).border = THIN
        ws_hol.cell(row, 3, f"=IF(A{row}=\"\",\"\",YEAR(A{row}))").border = THIN

    # Pre-format extra rows for future years (rows 21–80)
    for row in range(5 + len(HOLIDAYS_2026), 81):
        ws_hol.cell(row, 1).number_format = "YYYY-MM-DD"
        ws_hol.cell(row, 1).border = THIN
        ws_hol.cell(row, 1).fill = FILL_INPUT
        ws_hol.cell(row, 2).border = THIN
        ws_hol.cell(row, 2).fill = FILL_INPUT
        ws_hol.cell(row, 3, f'=IF(A{row}="","",YEAR(A{row}))').border = THIN

    ws_hol.column_dimensions["A"].width = 14
    ws_hol.column_dimensions["B"].width = 36
    ws_hol.column_dimensions["C"].width = 14

    # Define a workbook-scoped name for holiday dates range
    from openpyxl.workbook.defined_name import DefinedName

    wb.defined_names.add(DefinedName(name="HolidayDates", attr_text="'公眾假期'!$A$5:$A$80"))
    wb.defined_names.add(DefinedName(name="HolidayNames", attr_text="'公眾假期'!$B$5:$B$80"))

    # =====================================================================
    # Sheet: 請假登錄
    # =====================================================================
    ws_leave = wb.create_sheet("請假登錄")
    ws_leave["A1"] = "請假登錄（在對應人員／週份格輸入 AL 或 SL）"
    ws_leave["A1"].font = FONT_TITLE
    ws_leave.merge_cells("A1:H1")
    ws_leave["A2"] = (
        "此表欄位與「排班表」週份對齊。輸入 AL/SL 後，排班表對應格會自動覆蓋原班次。"
        "換年後若欄位仍留有舊假碼，請清除或改填新一年請假。"
    )
    ws_leave["A2"].font = FONT_NOTE
    ws_leave.merge_cells("A2:H2")

    # Week date headers — same formulas as main sheet (row 4)
    ws_leave["A4"] = "員工"
    ws_leave["A4"].fill = FILL_HEADER
    ws_leave["A4"].font = FONT_HDR
    ws_leave["A4"].border = THIN

    for w in range(NUM_WEEKS):
        col = 2 + w  # B=2
        cell = ws_leave.cell(4, col)
        cell.value = week_start_formula(w)
        cell.number_format = "M/D"
        cell.fill = FILL_HEADER
        cell.font = FONT_HDR
        cell.alignment = CENTER
        cell.border = THIN
        ws_leave.column_dimensions[get_column_letter(col)].width = 5

    for i in range(6):
        row = 5 + i
        # Name linked from 設定
        ws_leave.cell(row, 1, f"='設定'!B{EMP_FIRST_ROW+i}")
        ws_leave.cell(row, 1).font = FONT_BOLD
        ws_leave.cell(row, 1).border = THIN
        ws_leave.cell(row, 1).fill = FILL_SETTINGS
        for w in range(NUM_WEEKS):
            col = 2 + w
            cell = ws_leave.cell(row, col, None)
            cell.fill = FILL_INPUT
            cell.border = THIN
            cell.alignment = CENTER
            cell.font = FONT_BOLD

    dv_leave = DataValidation(type="list", formula1='"AL,SL"', allow_blank=True)
    ws_leave.add_data_validation(dv_leave)
    dv_leave.add(f"B5:{get_column_letter(1+NUM_WEEKS)}10")

    ws_leave.column_dimensions["A"].width = 14
    ws_leave.freeze_panes = "B5"

    # =====================================================================
    # Sheet: 排班表 (main)
    # =====================================================================
    ws = wb.create_sheet("排班表", 0)  # first sheet

    # --- Title & year echo ---
    ws["A1"] = "部門輪班更表（跨年模版）"
    ws["A1"].font = FONT_TITLE
    ws.merge_cells("A1:E1")

    ws["F1"] = "顯示年份"
    ws["F1"].font = FONT_BOLD
    ws["G1"] = f"={YEAR_CELL}"
    ws["G1"].font = Font(name="Microsoft JhengHei", bold=True, size=14, color="C00000")
    ws["G1"].fill = FILL_QUICK
    ws["G1"].border = THIN
    ws["G1"].alignment = CENTER
    ws["H1"] = "（修改請到「設定」工作表）"
    ws["H1"].font = FONT_SMALL

    # --- Quick view section (rows 3–10) ---
    ws["A3"] = "本週／下週快覽"
    ws["A3"].font = FONT_BOLD
    ws["A3"].fill = FILL_QUICK_HDR
    ws.merge_cells("A3:C3")

    ws["A4"] = "今天"
    ws["B4"] = "=TODAY()"
    ws["B4"].number_format = "YYYY-MM-DD"
    ws["B4"].font = FONT_BOLD

    # This week's Monday (WEEKDAY type 2: Mon=1)
    ws["A5"] = "本週起始（一）"
    ws["B5"] = "=B4-WEEKDAY(B4,2)+1"
    ws["B5"].number_format = "YYYY-MM-DD"
    ws["B5"].fill = FILL_QUICK
    ws["B5"].border = THIN

    ws["A6"] = "下週起始（一）"
    ws["B6"] = "=B5+7"
    ws["B6"].number_format = "YYYY-MM-DD"
    ws["B6"].fill = FILL_QUICK
    ws["B6"].border = THIN

    # MATCH column index of this week / next week in header row
    # Week headers will be in row 13, cols H onwards (col 8)
    header_row = 13
    name_row0 = 14  # first employee display row
    leave_sheet_row0 = 5

    first_col_letter = get_column_letter(FIRST_WEEK_COL)
    last_col_letter = get_column_letter(FIRST_WEEK_COL + NUM_WEEKS - 1)
    week_header_range = f"${first_col_letter}${header_row}:${last_col_letter}${header_row}"

    ws["D5"] = "本週欄位"
    ws["E5"] = f'=IFERROR(MATCH(B5,{week_header_range},0),"非本年")'
    ws["E5"].fill = FILL_QUICK
    ws["E5"].border = THIN

    ws["D6"] = "下週欄位"
    ws["E6"] = f'=IFERROR(MATCH(B6,{week_header_range},0),"非本年")'
    ws["E6"].fill = FILL_QUICK
    ws["E6"].border = THIN

    # Quick view table headers
    ws["A8"] = "人員"
    ws["B8"] = "本週班次"
    ws["C8"] = "下週班次"
    for col in range(1, 4):
        cell = ws.cell(8, col)
        cell.fill = FILL_QUICK_HDR
        cell.font = FONT_BOLD
        cell.border = THIN
        cell.alignment = CENTER

    for i in range(6):
        row = 9 + i
        # Name from settings
        ws.cell(row, 1, f"='設定'!B{EMP_FIRST_ROW+i}")
        ws.cell(row, 1).font = FONT_BOLD
        ws.cell(row, 1).border = THIN
        ws.cell(row, 1).fill = FILL_SETTINGS

        # INDEX into main roster row for this/next week
        roster_row = name_row0 + i
        # MATCH returns 1-based index within week range; INDEX from H{roster_row}
        ws.cell(
            row,
            2,
            f'=IF(OR($E$5="",$E$5="非本年"),"",'
            f'IFERROR(INDEX(${first_col_letter}{roster_row}:${last_col_letter}{roster_row},$E$5),""))',
        )
        ws.cell(
            row,
            3,
            f'=IF(OR($E$6="",$E$6="非本年"),"",'
            f'IFERROR(INDEX(${first_col_letter}{roster_row}:${last_col_letter}{roster_row},$E$6),""))',
        )
        for col in (2, 3):
            ws.cell(row, col).border = THIN
            ws.cell(row, col).alignment = CENTER
            ws.cell(row, col).font = FONT_BOLD
            ws.cell(row, col).fill = FILL_QUICK

    ws["D8"] = "圖例"
    ws["D8"].font = FONT_BOLD
    legend_quick = [("S", FILL_S), ("A", FILL_A), ("P", FILL_P), ("N", FILL_N), ("AL", FILL_AL), ("SL", FILL_SL)]
    for i, (code, fill) in enumerate(legend_quick):
        cell = ws.cell(8, 5 + i, code)
        cell.fill = fill
        cell.border = THIN
        cell.alignment = CENTER
        cell.font = FONT_BOLD

    ws["A15"] = ""  # placeholder — actual roster starts row 13

    # --- Stats + roster header row ---
    # Row 12: section title
    ws["A12"] = "全年排班主表（修改姓名／年份請到「設定」；請假請到「請假登錄」）"
    ws["A12"].font = FONT_BOLD
    ws.merge_cells("A12:G12")

    # Row 13: headers
    col_headers = [
        (1, "員工姓名"),
        (2, "已用年假\n(本年)"),
        (3, "剩餘年假\n(本年)"),
        (4, "已用病假\n(本年)"),
        (5, "累積補假\n結餘"),
        (6, "期初補假\n(可調)"),
        (7, "本年新增\n補假"),
    ]
    for col, text in col_headers:
        cell = ws.cell(header_row, col, text)
        cell.fill = FILL_HEADER
        cell.font = FONT_HDR
        cell.alignment = CENTER
        cell.border = THIN

    # Week date headers
    for w in range(NUM_WEEKS):
        col = FIRST_WEEK_COL + w
        cell = ws.cell(header_row, col)
        cell.value = week_start_formula(w)
        cell.number_format = "M/D"
        cell.fill = FILL_HEADER
        cell.font = FONT_HDR
        cell.alignment = CENTER
        cell.border = THIN
        ws.column_dimensions[get_column_letter(col)].width = 4.5

    # Employee rows 14–19
    for i, (name, role, offset) in enumerate(EMPLOYEES):
        row = name_row0 + i
        set_row = EMP_FIRST_ROW + i
        leave_row = leave_sheet_row0 + i

        # Name
        ws.cell(row, 1, f"='設定'!B{set_row}")
        ws.cell(row, 1).font = FONT_BOLD
        ws.cell(row, 1).border = THIN
        ws.cell(row, 1).alignment = CENTER
        ws.cell(row, 1).fill = FILL_SETTINGS

        # Stats formulas — only if name not blank
        # AL used: count AL in display row for weeks that have dates
        al_range = f"{first_col_letter}{row}:{last_col_letter}{row}"
        ws.cell(
            row,
            2,
            f'=IF(A{row}="","",COUNTIF({al_range},"AL"))',
        )
        ws.cell(
            row,
            3,
            f'=IF(A{row}="","",{QUOTA_CELL}-B{row})',
        )
        ws.cell(
            row,
            4,
            f'=IF(A{row}="","",COUNTIF({al_range},"SL"))',
        )

        # Opening compensatory balance (manual, cumulative across years)
        open_cell = ws.cell(row, 6, 0 if name else None)
        open_cell.fill = FILL_INPUT
        open_cell.border = THIN
        open_cell.alignment = CENTER

        # 本年新增補假 — uses helper row (holiday count per week) via SUMPRODUCT
        # Helper row number defined below as hol_count_row
        hol_count_row = name_row0 + 7  # row 21
        year_comp = (
            f'=IF(A{row}="","",'
            f'SUMPRODUCT('
            f'(({first_col_letter}{row}:{last_col_letter}{row}="A")'
            f'+({first_col_letter}{row}:{last_col_letter}{row}="P")'
            f'+({first_col_letter}{row}:{last_col_letter}{row}="N")'
            f'+({first_col_letter}{row}:{last_col_letter}{row}="S"))'
            f'*({first_col_letter}${hol_count_row}:{last_col_letter}${hol_count_row})))'
        )

        ws.cell(row, 7, year_comp)
        ws.cell(row, 7).border = THIN
        ws.cell(row, 7).fill = FILL_STAT
        ws.cell(row, 7).alignment = CENTER

        # Cumulative = opening + this year
        ws.cell(row, 5, f'=IF(A{row}="","",F{row}+G{row})')
        ws.cell(row, 5).border = THIN
        ws.cell(row, 5).fill = FILL_STAT
        ws.cell(row, 5).font = FONT_BOLD
        ws.cell(row, 5).alignment = CENTER

        for c in (2, 3, 4):
            ws.cell(row, c).border = THIN
            ws.cell(row, c).fill = FILL_STAT
            ws.cell(row, c).alignment = CENTER

        # Shift cells
        for w in range(NUM_WEEKS):
            col = FIRST_WEEK_COL + w
            col_l = get_column_letter(col)
            week_ref = f"{col_l}${header_row}"
            leave_ref = f"'請假登錄'!{get_column_letter(2+w)}{leave_row}"
            name_cell = f"'設定'!$B${set_row}"
            role_cell = f"'設定'!$C${set_row}"
            offset_cell = f"'設定'!$D${set_row}"

            # Only show if week header not blank
            formula = (
                f'=IF({week_ref}="","",'
                f'IF({name_cell}="","",'
                f'IF({leave_ref}<>"",{leave_ref},'
                f'IF({role_cell}="S","S",'
                f'IF({role_cell}="R",'
                f'IF(MOD(INT(({week_ref}-{BASE_CELL})/7)+IF({offset_cell}="",0,{offset_cell}),16)<4,"A",'
                f'IF(MOD(INT(({week_ref}-{BASE_CELL})/7)+IF({offset_cell}="",0,{offset_cell}),16)<8,"P",'
                f'IF(MOD(INT(({week_ref}-{BASE_CELL})/7)+IF({offset_cell}="",0,{offset_cell}),16)<12,"N","P"))),'
                f'"")))))'
            )
            cell = ws.cell(row, col, formula)
            cell.alignment = CENTER
            cell.border = THIN
            cell.font = FONT_BOLD

    # Holiday marker row under roster
    marker_row = name_row0 + 6  # row 20
    ws.cell(marker_row, 1, "公眾假期")
    ws.cell(marker_row, 1).font = FONT_SMALL
    ws.cell(marker_row, 1).fill = FILL_HOLIDAY
    ws.cell(marker_row, 1).border = THIN
    for w in range(NUM_WEEKS):
        col = FIRST_WEEK_COL + w
        col_l = get_column_letter(col)
        cell = ws.cell(marker_row, col)
        cell.value = (
            f'=IF({col_l}${header_row}="","",'
            f'IF(COUNTIFS(HolidayDates,">="&{col_l}${header_row},'
            f'HolidayDates,"<"&{col_l}${header_row}+7,HolidayDates,"<>")=0,"",'
            f'"假"&COUNTIFS(HolidayDates,">="&{col_l}${header_row},'
            f'HolidayDates,"<"&{col_l}${header_row}+7,HolidayDates,"<>")))'
        )
        cell.fill = FILL_HOLIDAY
        cell.font = FONT_SMALL
        cell.alignment = CENTER
        cell.border = THIN

    # Helper row: holiday count per week (used by 本年新增補假 SUMPRODUCT)
    hol_count_row = name_row0 + 7  # row 21
    ws.cell(hol_count_row, 1, "（系統）假期天數")
    ws.cell(hol_count_row, 1).font = FONT_SMALL
    ws.cell(hol_count_row, 1).fill = FILL_VACANT
    for w in range(NUM_WEEKS):
        col = FIRST_WEEK_COL + w
        col_l = get_column_letter(col)
        cell = ws.cell(
            hol_count_row,
            col,
            f'=IF({col_l}${header_row}="","",'
            f'COUNTIFS(HolidayDates,">="&{col_l}${header_row},'
            f'HolidayDates,"<"&{col_l}${header_row}+7,HolidayDates,"<>"))',
        )
        cell.font = FONT_SMALL
        cell.alignment = CENTER
        cell.border = THIN
        cell.fill = FILL_VACANT
    ws.row_dimensions[hol_count_row].hidden = True

    # Notes under table
    note_row = hol_count_row + 2
    ws.cell(note_row, 1, "使用提示")
    ws.cell(note_row, 1).font = FONT_BOLD
    tips = [
        "1. 黃色輸入格可改：設定區姓名／Offset／年假額度；排班表「期初補假」；「請假登錄」表對應週份格填 AL/SL。",
        "2. 換年：只改「設定」→顯示年份，並在「公眾假期」追加新假期。輪班基準日 2026-06-29 請勿改動。",
        "3. 衝突警示：同一週 AL+SL 合計超過 2 人時，該週欄位自動標紅。",
        "4. 補假：公眾假期當週若顯示 A/P/N/S（仍上班），計入本年新增補假；累積結餘＝期初＋本年（跨年請先把累積值寫入期初再換年）。",
        "5. 新增第六人：在設定填姓名、類型改 R、設 Offset（與其他人錯開 0/4/8/12），即可自動排班並可在請假登錄填假。",
        "6. 若公眾假期落在本年第一個週一之前（如 2026 元旦），該日不在主表週欄內，補假請手動調整「期初補假」。",
    ]
    for j, tip in enumerate(tips):
        ws.cell(note_row + 1 + j, 1, tip)
        ws.merge_cells(
            start_row=note_row + 1 + j,
            start_column=1,
            end_row=note_row + 1 + j,
            end_column=7,
        )
        ws.cell(note_row + 1 + j, 1).font = FONT_SMALL

    # Column widths
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 10
    ws.column_dimensions["F"].width = 10
    ws.column_dimensions["G"].width = 10

    # Row heights
    ws.row_dimensions[header_row].height = 32
    for r in range(name_row0, name_row0 + 6):
        ws.row_dimensions[r].height = 22

    # Freeze panes: freeze quick view + headers + name/stat cols
    # Freeze so row 13 (headers) and cols A-G stay visible when scrolling weeks
    ws.freeze_panes = "H14"

    # ----- Conditional formatting -----
    shift_range = f"{first_col_letter}{name_row0}:{last_col_letter}{name_row0+5}"
    quick_range = "B9:C14"

    def add_shift_cf(range_str):
        rules = [
            ("S", FILL_S),
            ("A", FILL_A),
            ("P", FILL_P),
            ("N", FILL_N),
            ("AL", FILL_AL),
            ("SL", FILL_SL),
        ]
        for code, fill in rules:
            ws.conditional_formatting.add(
                range_str,
                CellIsRule(operator="equal", formula=[f'"{code}"'], fill=fill),
            )

    add_shift_cf(shift_range)
    add_shift_cf("B9:C14")

    # Holiday week header highlighting
    for w in range(NUM_WEEKS):
        col_l = get_column_letter(FIRST_WEEK_COL + w)
        header_cell = f"{col_l}${header_row}"
        formula = (
            f'AND({header_cell}<>"",'
            f'COUNTIFS(HolidayDates,">="&{header_cell},'
            f'HolidayDates,"<"&{header_cell}+7,HolidayDates,"<>")>0)'
        )
        ws.conditional_formatting.add(
            f"{col_l}{header_row}",
            FormulaRule(formula=[formula], fill=FILL_HOLIDAY),
        )

    # Conflict: AL+SL count > 2 in same week column → red entire column (header + 6 people + marker)
    for w in range(NUM_WEEKS):
        col_l = get_column_letter(FIRST_WEEK_COL + w)
        col_range = f"{col_l}{header_row}:{col_l}{marker_row}"
        # Count AL and SL in employee rows
        formula = (
            f'(COUNTIF({col_l}${name_row0}:{col_l}${name_row0+5},"AL")+'
            f'COUNTIF({col_l}${name_row0}:{col_l}${name_row0+5},"SL"))>2'
        )
        ws.conditional_formatting.add(
            col_range,
            FormulaRule(formula=[formula], fill=FILL_WARN),
        )

    # =====================================================================
    # Sheet: 使用說明
    # =====================================================================
    ws_help = wb.create_sheet("使用說明")
    ws_help["A1"] = "部門更表模版 — 使用說明"
    ws_help["A1"].font = FONT_TITLE
    ws_help.merge_cells("A1:B1")

    sections = [
        (
            "一、每年年頭要做的兩件事",
            [
                "1. 打開「設定」工作表，把【顯示年份】改成新的年份（例如 2027）。",
                "   → 排班表所有週份日期（約 52/53 個星期一）會用 DATE／WEEKDAY 自動重算，無需複製新表。",
                "2. 打開「公眾假期」工作表，在清單下方空白列追加新一年的假期「日期」與「假期名稱」。",
                "   → 主表淺紫標示、補假計算會自動讀取新資料。",
                "3. 【輪班基準日】必須保持 2026-06-29 不變，否則全員輪班相位會錯位。",
                "4. 換年前建議：把各人「累積補假結餘」數字抄到「期初補假」，再改年份（本年新增會依新表重算）。",
                "5. 檢查「請假登錄」：換年後欄位對應新日期，請清除過期 AL/SL 或改填新一年請假。",
            ],
        ),
        (
            "二、人員流動（改姓名）",
            [
                "1. 只在「設定」工作表的六個姓名格修改；排班表、快覽區、請假登錄的姓名皆以公式連結，會自動更新。",
                "2. 班次類型：S＝固定正常班；R＝參與輪班；V＝空缺。",
                "3. 目前配置：ARTHUR＝S；PATRICK／CARREY／CESC／LUCAS＝R（Offset 0/4/8/12）；第六格＝空缺。",
            ],
        ),
        (
            "三、如何請年假／病假（AL／SL）",
            [
                "1. 到「請假登錄」工作表，在對應【人員 × 週份】格選擇或輸入 AL（年假）或 SL（病假）。",
                "2. 排班表同一位置會自動以 AL/SL 覆蓋原本的 S/A/P/N。",
                "3. 統計區「已用年假／已用病假」按顯示年份（本表週份）自動 COUNTIF；剩餘年假＝額度−已用。",
                "4. 同一週若 AL+SL 合計超過 2 人，該週整欄會自動標紅警示。",
            ],
        ),
        (
            "四、16 週輪班循環與 Offset（永久有效）",
            [
                "1. 錨點：輪班基準日 = 2026/6/29（星期一）。此日期永久固定，跨年也不改。",
                "2. 計算：週差 = INT((該週起始日 − 基準日)/7)；位置 = MOD(週差 + Offset, 16)。",
                "3. 循環段落：0–3 週→A；4–7 週→P；8–11 週→N；12–15 週→P；然後回到 A。",
                "4. Offset 0/4/8/12 錯開四人，保證任意一週 A、P、N 三班都有人當值（其中一段 P 會有兩人）。",
                "5. 因用「與基準日相差週數」計算，換年不會重置循環，班次連續銜接。",
            ],
        ),
        (
            "五、新增第六人",
            [
                "1. 在「設定」第 6 列填入姓名。",
                "2. 將「班次類型」由 V 改為 R（若要固定正常班則改 S）。",
                "3. 設定 Offset：建議選尚未使用、且能維持 A/P/N 覆蓋的值；若維持四人輪班結構，可與其中一人錯開討論後再定。",
                "4. 排班表該列公式已就緒，一有姓名即會顯示班次；並可在「請假登錄」為其登記 AL/SL。",
                "5. 「期初補假」可按實際結餘填入；統計欄會自動啟用。",
            ],
        ),
        (
            "六、顏色圖例",
            [
                "S＝白（正常班）｜A＝灰（早班）｜P＝綠（中班）｜N＝藍（晚班）",
                "AL＝黃（年假）｜SL＝橙（病假）｜公眾假期週＝表頭淺紫｜衝突週＝整欄紅底",
            ],
        ),
        (
            "七、工作表一覽",
            [
                "排班表 — 快覽＋全年主表＋統計（凍結首列區與姓名／統計欄）",
                "設定 — 年份、基準日、年假額度、六人姓名／類型／Offset",
                "請假登錄 — 輸入 AL/SL",
                "公眾假期 — 每年追加假期清單",
                "使用說明 — 本頁",
            ],
        ),
    ]

    r = 3
    for title, lines in sections:
        ws_help.cell(r, 1, title).font = FONT_BOLD
        ws_help.cell(r, 1).fill = FILL_SUBHDR
        ws_help.cell(r, 1).font = FONT_WHITE
        ws_help.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        r += 1
        for line in lines:
            ws_help.cell(r, 1, line).font = FONT_NORMAL
            ws_help.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
            r += 1
        r += 1

    ws_help.column_dimensions["A"].width = 100
    ws_help.column_dimensions["B"].width = 20

    # Reorder sheets
    order = ["排班表", "設定", "請假登錄", "公眾假期", "使用說明"]
    for idx, name in enumerate(order):
        wb.move_sheet(name, offset=idx - wb.sheetnames.index(name))

    out = "/workspace/部門更表模版_跨年.xlsx"
    wb.save(out)
    return out


if __name__ == "__main__":
    path = build_workbook()
    print(f"Created: {path}")
