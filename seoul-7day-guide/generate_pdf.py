#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""首爾七天遊 — 旅行社行程單格式 PDF"""

from pathlib import Path
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white, black
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle,
    PageBreak, KeepTogether, HRFlowable, Flowable,
)

ROOT = Path(__file__).resolve().parent
IMG = ROOT / "images_opt"
OUT = ROOT / "artifacts" / "首爾七天遊圖文攻略_2026.pdf"
OUT.parent.mkdir(parents=True, exist_ok=True)

FONT = "GuideFont"
for path in (
    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
    "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf",
):
    p = Path(path)
    if p.exists():
        try:
            pdfmetrics.registerFont(TTFont("GuideFont", str(p), subfontIndex=0))
            break
        except Exception:
            continue
else:
    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    FONT = "STSong-Light"

# Travel-agency document palette
NAVY = HexColor("#0B3D5C")
NAVY2 = HexColor("#145A7A")
GOLD = HexColor("#B8860B")
INK = HexColor("#222222")
MUTED = HexColor("#555555")
LINE = HexColor("#B8C0C8")
ROW_ALT = HexColor("#F3F6F8")
HEAD_BG = HexColor("#0B3D5C")
NOTE_BG = HexColor("#FFF8E7")
MEAL_BG = HexColor("#E8F1F5")
WHITE = white

PAGE_W, PAGE_H = A4
ML = MR = 12 * mm
MT = 14 * mm
MB = 14 * mm
CONTENT_W = PAGE_W - ML - MR  # ~186mm on A4 with 12mm margins


def P(text, size=9, leading=None, color=INK, align=TA_LEFT, name=None, **kw):
    return Paragraph(
        text,
        ParagraphStyle(
            name or f"p{size}{id(text)%9999}",
            fontName=FONT,
            fontSize=size,
            leading=leading or (size + 3.5),
            textColor=color,
            alignment=align,
            **kw,
        ),
    )


def img(name, w=CONTENT_W, h=42 * mm):
    stem = Path(name).stem
    path = IMG / f"{stem}.jpg"
    if not path.exists():
        path = IMG / name
    if not path.exists():
        return Spacer(1, 1)
    return Image(str(path), width=w, height=h, kind="proportional")


def hline():
    return HRFlowable(width="100%", thickness=0.8, color=NAVY, spaceBefore=2, spaceAfter=4)


def thinline():
    return HRFlowable(width="100%", thickness=0.4, color=LINE, spaceBefore=2, spaceAfter=4)


def section_title(text):
    t = Table([[P(f"<b>{text}</b>", 10, 13, WHITE)]], colWidths=[CONTENT_W])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), NAVY),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return t


def kv_table(rows, col1=42 * mm, col2=None):
    col2 = col2 or (CONTENT_W - col1)
    data = []
    for k, v in rows:
        data.append([P(f"<b>{k}</b>", 8.5, 11.5, NAVY), P(v, 8.5, 11.5, INK)])
    t = Table(data, colWidths=[col1, col2])
    style = [
        ("GRID", (0, 0), (-1, -1), 0.4, LINE),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
        ("BACKGROUND", (0, 0), (0, -1), ROW_ALT),
    ]
    t.setStyle(TableStyle(style))
    return t


def schedule_table(rows):
    """rows: list of (time, content, meal, note)"""
    header = [
        P("<b>時間</b>", 8, 10, WHITE, TA_CENTER),
        P("<b>行程內容</b>", 8, 10, WHITE, TA_CENTER),
        P("<b>餐食</b>", 8, 10, WHITE, TA_CENTER),
        P("<b>交通／備註</b>", 8, 10, WHITE, TA_CENTER),
    ]
    data = [header]
    for time, content, meal, note in rows:
        data.append([
            P(time, 8, 11, INK, TA_CENTER),
            P(content, 8, 11, INK),
            P(meal, 8, 11, INK, TA_CENTER),
            P(note, 7.5, 10.5, MUTED),
        ])
    # time | content | meal | note
    t = Table(data, colWidths=[28 * mm, 88 * mm, 18 * mm, CONTENT_W - 134 * mm])
    cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), HEAD_BG),
        ("GRID", (0, 0), (-1, -1), 0.4, LINE),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3.5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3.5),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
        ("BACKGROUND", (2, 1), (2, -1), MEAL_BG),
    ]
    for i in range(1, len(data)):
        if i % 2 == 0:
            cmds.append(("BACKGROUND", (0, i), (1, i), ROW_ALT))
            cmds.append(("BACKGROUND", (3, i), (3, i), ROW_ALT))
    t.setStyle(TableStyle(cmds))
    return t


def note_box(title, text):
    data = [
        [P(f"<b>■ {title}</b>", 8.5, 11, NAVY)],
        [P(text, 8, 11, INK)],
    ]
    t = Table(data, colWidths=[CONTENT_W])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), NOTE_BG),
        ("BOX", (0, 0), (-1, -1), 0.6, GOLD),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (0, 0), 5),
        ("BOTTOMPADDING", (0, -1), (0, -1), 5),
        ("TOPPADDING", (0, 1), (0, 1), 1),
    ]))
    return t


def day_header(day_no, date, weekday, theme):
    left = P(f"<b>DAY {day_no}</b>", 14, 18, WHITE, TA_CENTER)
    right = P(f"<b>{date}（{weekday}）</b>　{theme}", 10, 14, WHITE)
    t = Table([[left, right]], colWidths=[28 * mm, CONTENT_W - 28 * mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), GOLD),
        ("BACKGROUND", (1, 0), (1, 0), NAVY),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    return t


def footer(canvas, doc):
    canvas.saveState()
    canvas.setStrokeColor(NAVY)
    canvas.setLineWidth(1.2)
    canvas.line(ML, PAGE_H - 8 * mm, PAGE_W - MR, PAGE_H - 8 * mm)
    canvas.setFillColor(NAVY)
    canvas.setFont(FONT, 7)
    canvas.drawString(ML, PAGE_H - 6.5 * mm, "首爾七天休閒度假行程單｜ITINERARY")
    canvas.drawRightString(PAGE_W - MR, PAGE_H - 6.5 * mm, "CONFIDENTIAL / FOR TRAVELER USE")
    canvas.setStrokeColor(LINE)
    canvas.setLineWidth(0.5)
    canvas.line(ML, 10 * mm, PAGE_W - MR, 10 * mm)
    canvas.setFillColor(MUTED)
    canvas.setFont(FONT, 7)
    canvas.drawString(ML, 6.5 * mm, "Hotel POCO Seongsu｜2026.08.08–08.14")
    canvas.drawRightString(PAGE_W - MR, 6.5 * mm, f"第 {doc.page} 頁")
    canvas.restoreState()


# ─── Pages ───────────────────────────────────────────────

def page_cover(story):
    # Agency-style masthead
    mast = Table(
        [[P("<b>TRAVEL ITINERARY</b>", 9, 12, GOLD, TA_CENTER)],
         [P("<b>首爾七天六夜　休閒度假行程單</b>", 18, 24, WHITE, TA_CENTER)],
         [P("SEOUL 7 DAYS / 6 NIGHTS　｜　LEISURE PACE", 9, 12, HexColor("#C8D9E4"), TA_CENTER)]],
        colWidths=[CONTENT_W],
    )
    mast.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), NAVY),
        ("TOPPADDING", (0, 0), (0, 0), 8),
        ("BOTTOMPADDING", (0, -1), (0, -1), 8),
        ("TOPPADDING", (0, 1), (0, 1), 2),
        ("BOTTOMPADDING", (0, 1), (0, 1), 2),
    ]))
    story.append(mast)
    story.append(Spacer(1, 3 * mm))
    story.append(img("cover-seoul-summer.jpg", CONTENT_W, 48 * mm))
    story.append(Spacer(1, 3 * mm))

    story.append(section_title("一、旅客與行程基本資料"))
    story.append(Spacer(1, 2 * mm))
    story.append(kv_table([
        ("行程名稱", "首爾聖水慢活七天遊（圖文執行版）"),
        ("旅遊日期", "2026年8月8日（六）至 2026年8月14日（五）｜共7天6夜"),
        ("旅客人數", "兩位成人"),
        ("行程風格", "休閒度假為主；每日2–3個重點，保留咖啡／漫步時間"),
        ("交通偏好", "地鐵為主＋1公里內步行；盡量減少打車"),
        ("季節提醒", "8月首爾盛夏：防曬、補水；室內冷氣強請備薄外套"),
    ]))
    story.append(Spacer(1, 3 * mm))

    story.append(section_title("二、住宿資料"))
    story.append(Spacer(1, 2 * mm))
    story.append(kv_table([
        ("飯店名稱", "Hotel POCO Seongsu（호텔 포코 성수）"),
        ("地址", "서울 성동구 성수이로 96（Seongsu-ro 96, Seongdong-gu）"),
        ("交通", "地鐵2號線　聖水站（성수）3號出口　右轉步行約50m（約1分鐘）"),
        ("電話", "預約 02-3677-6676　／　前台 02-462-9610"),
        ("入住／退房", "Day1 下午入住　｜　Day7 建議 12:00–12:30 前退房出發"),
    ]))
    story.append(Spacer(1, 3 * mm))

    story.append(section_title("三、航班資料（請以電子機票為準）"))
    story.append(Spacer(1, 2 * mm))
    flight_h = [
        P("<b>航段</b>", 8, 10, WHITE, TA_CENTER),
        P("<b>航班</b>", 8, 10, WHITE, TA_CENTER),
        P("<b>日期</b>", 8, 10, WHITE, TA_CENTER),
        P("<b>起飛</b>", 8, 10, WHITE, TA_CENTER),
        P("<b>抵達</b>", 8, 10, WHITE, TA_CENTER),
        P("<b>備註</b>", 8, 10, WHITE, TA_CENTER),
    ]
    flight_rows = [flight_h,
        [P("去程 HKG→ICN", 8, 10), P("<b>UO630</b>", 8, 10, INK, TA_CENTER),
         P("8/8（六）", 8, 10, INK, TA_CENTER), P("約 09:45", 8, 10, INK, TA_CENTER),
         P("約 14:25", 8, 10, INK, TA_CENTER), P("預估16:00–17:00抵飯店", 7.5, 10)],
        [P("回程 ICN→HKG", 8, 10), P("<b>UO631</b>", 8, 10, INK, TA_CENTER),
         P("8/14（五）", 8, 10, INK, TA_CENTER), P("約15:25–16:10", 8, 10, INK, TA_CENTER),
         P("約18:15–19:00", 8, 10, INK, TA_CENTER), P("建議12:00前出發", 7.5, 10)],
    ]
    ft = Table(flight_rows, colWidths=[36*mm, 22*mm, 24*mm, 28*mm, 28*mm, CONTENT_W-138*mm])
    ft.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), HEAD_BG),
        ("GRID", (0, 0), (-1, -1), 0.4, LINE),
        ("BACKGROUND", (0, 2), (-1, 2), ROW_ALT),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(ft)
    story.append(Spacer(1, 3 * mm))

    story.append(section_title("四、固定行程（不可更改）"))
    story.append(Spacer(1, 2 * mm))
    story.append(kv_table([
        ("8/9（日）18:00", "Alla Prima 生日晚餐（已訂位）｜학동로17길 13｜套餐約19–31萬韓元／人｜下午留白化妝穿搭"),
        ("8/12（三）", "MORE ON HAIR 江南店美髮（강남대로475, 3F，非江南總店）＋유후네일 신논현美甲銜接"),
    ]))
    story.append(PageBreak())


def page_overview(story):
    story.append(section_title("五、七日行程總覽表"))
    story.append(Spacer(1, 2 * mm))
    h = [P(f"<b>{x}</b>", 7.5, 10, WHITE, TA_CENTER) for x in
         ("日序", "日期", "主題", "上午／午間", "下午／晚上", "住宿")]
    rows = [h,
        [P("Day1", 7.5, 10, INK, TA_CENTER), P("8/8 六", 7.5, 10, INK, TA_CENTER),
         P("抵達慢活", 7.5, 10), P("UO630抵韓／入住", 7.5, 10),
         P("大林倉庫・성수연방・輕晚餐", 7.5, 10), P("POCO", 7.5, 10, INK, TA_CENTER)],
        [P("Day2", 7.5, 10, INK, TA_CENTER), P("8/9 日", 7.5, 10, INK, TA_CENTER),
         P("麵包＋慶生", 7.5, 10), P("首爾林・Cafe Onion", 7.5, 10),
         P("飯店梳化・Alla Prima 18:00", 7.5, 10), P("POCO", 7.5, 10, INK, TA_CENTER)],
        [P("Day3", 7.5, 10, INK, TA_CENTER), P("8/10 一", 7.5, 10, INK, TA_CENTER),
         P("韓牛＋DDP＋雞", 7.5, 10), P("인생한우（馬場）", 7.5, 10),
         P("DDP・진옥화晚餐", 7.5, 10), P("POCO", 7.5, 10, INK, TA_CENTER)],
        [P("Day4", 7.5, 10, INK, TA_CENTER), P("8/11 二", 7.5, 10, INK, TA_CENTER),
         P("設計＋咖啡", 7.5, 10), P("안목湯飯・Grey Penguin", 7.5, 10),
         P("29CM／LCDC／Object・咖啡", 7.5, 10), P("POCO", 7.5, 10, INK, TA_CENTER)],
        [P("Day5", 7.5, 10, INK, TA_CENTER), P("8/12 三", 7.5, 10, INK, TA_CENTER),
         P("美髮美甲", 7.5, 10), P("유후네일美甲", 7.5, 10),
         P("MORE ON HAIR 江南店", 7.5, 10), P("POCO", 7.5, 10, INK, TA_CENTER)],
        [P("Day6", 7.5, 10, INK, TA_CENTER), P("8/13 四", 7.5, 10, INK, TA_CENTER),
         P("弘大慢逛", 7.5, 10), P("弘大早午餐・逛街", 7.5, 10),
         P("延南森林路・炸雞夜", 7.5, 10), P("POCO", 7.5, 10, INK, TA_CENTER)],
        [P("Day7", 7.5, 10, INK, TA_CENTER), P("8/14 五", 7.5, 10, INK, TA_CENTER),
         P("退房返港", 7.5, 10), P("早餐・最後採購", 7.5, 10),
         P("12:00前出發・UO631", 7.5, 10), P("—", 7.5, 10, INK, TA_CENTER)],
    ]
    t = Table(rows, colWidths=[14*mm, 18*mm, 24*mm, 42*mm, 50*mm, 16*mm])
    cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), HEAD_BG),
        ("GRID", (0, 0), (-1, -1), 0.35, LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2.5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2.5),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
    ]
    for i in range(2, len(rows), 2):
        cmds.append(("BACKGROUND", (0, i), (-1, i), ROW_ALT))
    t.setStyle(TableStyle(cmds))
    story.append(t)
    story.append(Spacer(1, 4 * mm))

    story.append(section_title("六、費用說明（自行消費估算・兩人）"))
    story.append(Spacer(1, 2 * mm))
    story.append(P(
        "本行程為「自助執行行程單」，以下為餐飲／交通／體驗粗估，不含機票與飯店房費。"
        "Alla Prima、美髮美甲為旅客既定消費，金額以現場／預約單為準。",
        8.5, 12, MUTED,
    ))
    story.append(Spacer(1, 2 * mm))
    bh = [P(f"<b>{x}</b>", 7.5, 10, WHITE, TA_CENTER) for x in ("日序", "餐飲粗估", "交通", "其他", "合計約")]
    brows = [bh,
        [P("Day1", 7.5, 10, INK, TA_CENTER), P("4–8萬", 7.5, 10, INK, TA_CENTER), P("機場1–1.6萬", 7.5, 10, INK, TA_CENTER), P("雜費1萬", 7.5, 10, INK, TA_CENTER), P("6–10萬", 7.5, 10, INK, TA_CENTER)],
        [P("Day2", 7.5, 10, INK, TA_CENTER), P("Onion4–6萬＋晚餐45–75萬", 7.5, 10, INK, TA_CENTER), P("地鐵0.5萬", 7.5, 10, INK, TA_CENTER), P("—", 7.5, 10, INK, TA_CENTER), P("50–82萬", 7.5, 10, INK, TA_CENTER)],
        [P("Day3", 7.5, 10, INK, TA_CENTER), P("韓牛12–25萬＋雞4–6萬", 7.5, 10, INK, TA_CENTER), P("地鐵0.6萬", 7.5, 10, INK, TA_CENTER), P("DDP1–3萬", 7.5, 10, INK, TA_CENTER), P("17–35萬", 7.5, 10, INK, TA_CENTER)],
        [P("Day4", 7.5, 10, INK, TA_CENTER), P("안목6–8萬＋咖啡4–7萬", 7.5, 10, INK, TA_CENTER), P("少", 7.5, 10, INK, TA_CENTER), P("小物0–15萬", 7.5, 10, INK, TA_CENTER), P("10–30萬", 7.5, 10, INK, TA_CENTER)],
        [P("Day5", 7.5, 10, INK, TA_CENTER), P("餐3–5萬", 7.5, 10, INK, TA_CENTER), P("地鐵0.5萬", 7.5, 10, INK, TA_CENTER), P("美甲4–10萬＋美髮另計", 7.5, 10, INK, TA_CENTER), P("視服務項目", 7.5, 10, INK, TA_CENTER)],
        [P("Day6", 7.5, 10, INK, TA_CENTER), P("餐＋炸雞9–15萬", 7.5, 10, INK, TA_CENTER), P("地鐵0.5萬", 7.5, 10, INK, TA_CENTER), P("小物2–8萬", 7.5, 10, INK, TA_CENTER), P("11–23萬", 7.5, 10, INK, TA_CENTER)],
        [P("Day7", 7.5, 10, INK, TA_CENTER), P("早餐2–4萬", 7.5, 10, INK, TA_CENTER), P("機場1–1.6萬", 7.5, 10, INK, TA_CENTER), P("—", 7.5, 10, INK, TA_CENTER), P("3–9萬", 7.5, 10, INK, TA_CENTER)],
    ]
    bt = Table(brows, colWidths=[16*mm, 55*mm, 32*mm, 40*mm, 21*mm])
    bcmds = [
        ("BACKGROUND", (0, 0), (-1, 0), HEAD_BG),
        ("GRID", (0, 0), (-1, -1), 0.35, LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]
    for i in range(2, len(brows), 2):
        bcmds.append(("BACKGROUND", (0, i), (-1, i), ROW_ALT))
    bt.setStyle(TableStyle(bcmds))
    story.append(bt)
    story.append(Spacer(1, 3 * mm))
    story.append(note_box(
        "匯率與付款",
        "匯率請以出行前銀行／Google為準（常見約1 HKD≈170–185 KRW，會波動）。"
        "信用卡普及；建議備5–10萬韓元現金。單位：韓元。",
    ))
    story.append(PageBreak())


def add_day(story, day_no, date, weekday, theme, photo, rows, budget_text, alt_title, alt_text, stay="Hotel POCO Seongsu"):
    blocks = [
        day_header(day_no, date, weekday, theme),
        Spacer(1, 2 * mm),
    ]
    if photo:
        blocks += [img(photo, CONTENT_W, 40 * mm), Spacer(1, 2 * mm)]
    blocks += [
        P(f"<b>住宿</b>：{stay}", 8, 11, MUTED),
        Spacer(1, 1.5 * mm),
        schedule_table(rows),
        Spacer(1, 2.5 * mm),
        P(f"<b>當日預算（兩人・約）</b>：{budget_text}", 8.5, 12, NAVY),
        Spacer(1, 2 * mm),
        note_box(alt_title, alt_text),
    ]
    story.append(KeepTogether(blocks))
    story.append(PageBreak())


def pages_days(story):
    add_day(
        story, 1, "2026/08/08", "六", "抵達聖水・落地慢活",
        "day1-daelim-cafe.jpg",
        [
            ("09:45–14:25", "<b>航班 UO630</b>　香港 HKG → 仁川 ICN（實際時刻以電子機票為準）", "機上", "入境＋領行李約45–75分"),
            ("15:30–17:00", "<b>機場→飯店</b>　AREX→弘大入口轉2號線→聖水站3號出口步行約1分", "—", "全程約70–90分；預估16–17時抵店"),
            ("17:00–18:00", "<b>Check-in</b>　Hotel POCO Seongsu　沖澡、調冷氣、確認地圖App", "—", "성수이로 96"),
            ("18:00–19:30", "<b>周邊散步</b>　①대림창고（성수이로78）②Musinsa＠대림창고（성수이로74）③성수연방／천상가옥（성수이로14길14）", "咖啡可", "飯店同路，步行為主"),
            ("19:30–21:00", "<b>輕鬆晚餐</b>　A.꿉당성수（성수이로20길10）炭火豬頸肉　B.飯店5分內簡餐", "晚餐", "週末꿉당可能排隊"),
            ("21:00後", "返回飯店休息，早睡養精神", "—", "不排景點"),
        ],
        "6–10萬韓元（含機場交通）",
        "備選方案",
        "抵達過晚改外帶回房。大林倉庫人多改坐성수연방吹冷氣。AREX人潮可改機場巴士後轉地鐵。",
    )

    add_day(
        story, 2, "2026/08/09", "日", "Cafe Onion 麵包午餐・Alla Prima 慶生晚餐",
        "day2-onion-bakery.jpg",
        [
            ("09:30–11:00", "<b>首爾林（서울숲）</b>　樹蔭步道45–60分（趁涼，不排滿園）", "—", "聖水步行或盆唐線首爾林站"),
            ("11:30–13:30", "<b>午餐：Cafe Onion 성수</b>　아차산로9길8｜必點설산팡도르、앙버터＋冰美式｜先佔座再買麵包", "午餐", "聖水2號出口步行2–3分｜人均1.5–2.5萬"),
            ("13:30–16:00", "<b>回飯店休息</b>　午睡、保濕、化妝穿搭（smart casual；勿短褲拖鞋）", "—", "為18:00慶生留白"),
            ("16:30–17:50", "<b>前往鶴洞</b>　聖水→建大入口轉7號線→鶴洞站6號出口｜可選더스퀘어咖啡補妝", "—", "車程約30–40分｜提前抵達"),
            ("18:00–21:00", "<b>【固定】Alla Prima 生日晚餐</b>　학동로17길13｜套餐約19–31萬／人｜約2.5–3小時", "晚餐", "遲到約15分可能no-show｜電話02-511-2555"),
            ("21:00後", "7號線→建大入口→2號線返回聖水飯店", "—", "可購氣泡水慶祝收工"),
        ],
        "50–82萬韓元（含Alla Prima）",
        "備選方案",
        "下雨取消首爾林，改Onion長坐或성수연방。Onion大排長龍可外帶，或改밀도성수（왕십리로96）吐司外帶。",
    )

    add_day(
        story, 3, "2026/08/10", "一", "馬場韓牛 → DDP → 一隻雞晚餐",
        "day3-ddp.jpg",
        [
            ("11:00–11:30", "<b>前往馬場洞</b>　聖水→往十里轉5號線→馬場站", "—", "約20–25分"),
            ("11:30–14:00", "<b>午餐：인생한우</b>　마장로31길43 2層｜選1++꽃등심／살치살／채끝｜先點600–800g", "午餐", "上桌費홀約4,000｜兩人約12–25萬"),
            ("14:30–17:30", "<b>東大門設計廣場 DDP</b>　室內曲線建築＋當期展覽（查ddp.or.kr）｜盛夏避暑", "咖啡可", "馬場／往十里轉東大門歷史文化公園站約15–20分"),
            ("18:00–20:00", "<b>晚餐：진옥화할매원조닭한마리</b>　종로40가길18｜不預約現場排隊｜雙人一隻雞＋年糕／粉條", "晚餐", "東大門站9號出口步行6–8分｜約4–6萬"),
            ("20:30後", "地鐵返回聖水飯店休息", "—", "明天設計家具日"),
        ],
        "17–35萬韓元",
        "備選方案",
        "DDP人多可改Doota／現代City Mall吹冷氣。一隻雞排隊過長：同巷其他店，或回聖水남다른본가닭한마리（연무장7길12 2層）。",
    )

    add_day(
        story, 4, "2026/08/11", "二", "安木湯飯＋聖水設計家具＋咖啡串連",
        "day4-design-shop.jpg",
        [
            ("11:30–13:00", "<b>午餐：안목 성수</b>　뚝섬로13길34｜豬國飯＋冰冷모듬수육｜可用CatchTable候位", "午餐", "聖水3號出口步行7–10分｜人均約3–4萬"),
            ("13:30–14:30", "<b>Cafe① Grey Penguin</b>　서울숲4길26-14｜塔派＋冰飲，飯後休息", "甜點", "近首爾林"),
            ("14:30–17:30", "<b>設計店動線</b>　①29CM HOME（연무장길57）②29CM HOME2（110）③LCDC（연무장17길10）④Object（서울숲길36 2層）", "—", "Object約12:30–20:30｜週一休故排週二"),
            ("彈性插入", "<b>Cafe②</b>　대림창고（성수이로78）或성수연방・천상가옥（성수이로14길14 3層）", "咖啡", "一天認真坐2間即可"),
            ("可選", "<b>아모레성수</b>　아차산로11길7｜10:30–20:30（週一休→今日可）試香吹冷氣", "—", "聖水2號出口步行約5分"),
            ("晚上", "輕食回飯店；明天美容日少油炸", "輕食", "—"),
        ],
        "10–30萬韓元（含小物則上限上調）",
        "備選方案",
        "店休／下雨縮成兩間設計店＋Grey Penguin與천상가옥長坐，不硬走연무장全線。",
    )

    add_day(
        story, 5, "2026/08/12", "三", "美髮＋美甲日（新論峴一站搞定）",
        "day5-beauty-salon.jpg",
        [
            ("10:00–10:40", "<b>移動</b>　聖水2號線直達新論峴站，2號出口出站", "—", "約20–25分"),
            ("11:00–13:00", "<b>【固定】美甲：유후네일 신논현</b>　봉은사로1길37｜Naver搜「유후네일 신논현」預約", "—", "步行約5–8分｜凝膠約4–7萬起"),
            ("13:00–14:00", "<b>輕午餐</b>　江南大路商圈簡餐（少沾手食物）", "午餐", "人均1.2–2萬"),
            ("14:00–18:00", "<b>【固定】美髮：MORE ON HAIR 江南店</b>　강남대로475, 3F（非江南總店）｜新論峴2號出口步行", "—", "以Naver預約單為準｜週一週二公休"),
            ("晚上", "返回聖水輕食／蛋糕慶祝新造型", "輕食", "早休息"),
        ],
        "餐3–5萬＋美甲4–10萬＋美髮（剪染燙常見10–40萬＋，依項目）",
        "備選方案",
        "美甲約滿：Naver「신논현역 네일샵」篩預約可；或미쥬네일（봉은사로4길23一帶）。優先保住美髮時段。",
    )

    add_day(
        story, 6, "2026/08/13", "四", "弘大慢逛日（延南森林路＋炸雞夜）",
        "day6-hongdae.jpg",
        [
            ("10:30–11:30", "<b>前往弘大</b>　聖水2號線直達弘大入口站", "—", "約30分"),
            ("11:30–13:00", "<b>早午餐＋咖啡</b>　弘大商圈側街選店（避開最擠巷口）", "午餐", "人均1.5–2.5萬"),
            ("13:00–17:30", "<b>弘大→延南洞</b>　①商圈服飾彩妝選品（多進店躲暑）②延南洞③京義線森林路延南段散步＋咖啡45–60分", "咖啡", "盛夏主街曬，傍晚走森林路較舒適"),
            ("18:30–20:30", "<b>炸雞夜（留弘大）</b>　Naver搜「홍대 치킨」依評分距離選店＋生啤", "晚餐", "兩人約4–7萬｜減少折返"),
            ("晚上", "返回聖水收拾明日托運行李，確認UO631航廈", "—", "備援炸雞：레츠잇치킨／교촌성수역점"),
        ],
        "11–23萬韓元",
        "備選方案",
        "暴晒／雷雨：縮短逛街，改延南咖啡長坐；大雨改室內選品，炸雞可外帶回飯店。",
    )

    add_day(
        story, 7, "2026/08/14", "五", "退房返港（UO631）",
        "day7-icn-airport.jpg",
        [
            ("08:30–10:00", "<b>收拾行李＋早餐</b>　Onion外帶／便利店／천상가옥輕食", "早餐", "不排景點"),
            ("10:00–11:30", "<b>最後採購（可選）</b>　僅限步行10分內藥妝或大林倉庫／Musinsa", "—", "液體刀剪注意托運規定"),
            ("12:00–12:30前", "<b>退房出發</b>　聖水→弘大入口轉AREX→仁川機場（約70–90分）＋提前2小時候機", "—", "務必預留緩衝"),
            ("13:30–起飛前", "<b>機場程序</b>　報到、托運、安檢、退稅；航廈內可當最後一餐", "可", "確認航廈（T1/T2）"),
            ("約15:25–16:10", "<b>航班 UO631</b>　ICN → HKG｜抵達香港約18:15–19:00（以機票為準）", "機上", "起飛前再確認航班狀態"),
        ],
        "3–9萬韓元（含機場交通）",
        "備選方案",
        "AREX異常改機場巴士或計程車。建議起飛前3小時再確認航廈與航班。",
        stay="当日退房",
    )


def page_prep(story):
    story.append(section_title("七、行前準備與注意事項"))
    story.append(Spacer(1, 2 * mm))
    story.append(kv_table([
        ("證件", "護照效期、韓國入境規定、電子機票、飯店與餐廳訂位截圖"),
        ("上網", "機場eSIM／旅遊SIM（5–7天）；可Hotspot共用"),
        ("必備App", "Naver Map（找路／預約）、Papago（翻譯）、CatchTable（候位）、Kakao T（叫車）、航空App"),
        ("盛夏裝備", "SPF50+、帽子、摺疊傘、補水壺、薄外套（室內冷氣）、舒適走路鞋"),
        ("付款", "信用卡為主＋現金5–10萬韓元；部分美甲店可能偏好轉帳／現金"),
        ("預約再確認", "出發前48小時：Alla Prima 18:00、MORE ON HAIR、유후네일、UO630/631航廈"),
    ]))
    story.append(Spacer(1, 3 * mm))
    story.append(section_title("八、美甲預約指引（可截圖）"))
    story.append(Spacer(1, 2 * mm))
    story.append(kv_table([
        ("店名", "유후네일 신논현점（YouWho Nail Sinnonhyeon）"),
        ("地址", "서울 강남구 봉은사로1길 37"),
        ("搜尋", "Naver Map →「유후네일 신논현」→ N예약"),
        ("銜接", "新論峴2號出口商圈　↔　MORE ON HAIR（강남대로475, 3F）步行可達"),
        ("建議時序", "11:00美甲 → 午餐 → 14:00美髮（可依預約單微調）"),
    ]))
    story.append(Spacer(1, 4 * mm))
    end = Table(
        [[P("<b>祝旅途愉快　Have a wonderful trip in Seoul</b>", 11, 15, WHITE, TA_CENTER)],
         [P("本行程單為自助執行參考；店家營業、價格、公休日與航班時刻請以當日官方資訊為準。", 8, 11, HexColor("#C8D9E4"), TA_CENTER)]],
        colWidths=[CONTENT_W],
    )
    end.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), NAVY),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(end)


def main():
    doc = SimpleDocTemplate(
        str(OUT), pagesize=A4,
        leftMargin=ML, rightMargin=MR, topMargin=MT, bottomMargin=MB,
        title="首爾七天休閒度假行程單 2026.08.08–08.14",
        author="Seoul Travel Itinerary",
    )
    story = []
    page_cover(story)
    page_overview(story)
    pages_days(story)
    page_prep(story)
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    main()
