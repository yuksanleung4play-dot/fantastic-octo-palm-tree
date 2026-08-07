#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""首爾七天遊圖文攻略 PDF（2026.8.8–8.14 定稿版）"""

from pathlib import Path
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle,
    PageBreak, HRFlowable,
)

ROOT = Path(__file__).resolve().parent
IMG = ROOT / "images_opt"
OUT = ROOT / "artifacts" / "首爾七天遊圖文攻略_2026.pdf"
OUT.parent.mkdir(parents=True, exist_ok=True)

FONT_REG = "GuideFont"
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
    FONT_REG = "STSong-Light"

C_INK = HexColor("#1C2421")
C_MUTED = HexColor("#5A655F")
C_ACCENT = HexColor("#0E6B5C")
C_ACCENT2 = HexColor("#C45C26")
C_SOFT = HexColor("#E6EFEA")
C_LINE = HexColor("#D5D0C8")
C_TIP = HexColor("#FFF6E8")

PAGE_W, PAGE_H = A4
MARGIN = 15 * mm


def styles():
    s = {}
    s["h1"] = ParagraphStyle(
        "h1", fontName=FONT_REG, fontSize=16, leading=22,
        textColor=C_ACCENT, spaceBefore=2, spaceAfter=6,
    )
    s["h2"] = ParagraphStyle(
        "h2", fontName=FONT_REG, fontSize=12, leading=16,
        textColor=C_INK, spaceBefore=8, spaceAfter=3,
    )
    s["h3"] = ParagraphStyle(
        "h3", fontName=FONT_REG, fontSize=10.5, leading=14,
        textColor=C_ACCENT2, spaceBefore=5, spaceAfter=2,
    )
    s["body"] = ParagraphStyle(
        "body", fontName=FONT_REG, fontSize=9.2, leading=13.8,
        textColor=C_INK, alignment=TA_JUSTIFY, spaceAfter=3,
    )
    s["meta"] = ParagraphStyle(
        "meta", fontName=FONT_REG, fontSize=8.8, leading=12.5,
        textColor=C_MUTED, spaceAfter=2,
    )
    s["caption"] = ParagraphStyle(
        "caption", fontName=FONT_REG, fontSize=7.8, leading=10.5,
        textColor=C_MUTED, alignment=TA_CENTER, spaceBefore=1, spaceAfter=6,
    )
    s["bullet"] = ParagraphStyle(
        "bullet", fontName=FONT_REG, fontSize=9.2, leading=13.2,
        textColor=C_INK, leftIndent=6, spaceAfter=1.5,
    )
    s["tip"] = ParagraphStyle(
        "tip", fontName=FONT_REG, fontSize=8.6, leading=12.2,
        textColor=C_INK, spaceAfter=1,
    )
    s["table"] = ParagraphStyle(
        "table", fontName=FONT_REG, fontSize=7.8, leading=10.5,
        textColor=C_INK, alignment=TA_LEFT,
    )
    s["table_h"] = ParagraphStyle(
        "table_h", fontName=FONT_REG, fontSize=7.8, leading=10.5,
        textColor=white, alignment=TA_CENTER,
    )
    return s


S = styles()


def img(name, w=172 * mm, h=68 * mm):
    stem = Path(name).stem
    path = IMG / f"{stem}.jpg"
    if not path.exists():
        path = IMG / name
    if not path.exists():
        return Spacer(1, 1)
    return Image(str(path), width=w, height=h, kind="proportional")


def hr():
    return HRFlowable(width="100%", thickness=0.5, color=C_LINE, spaceBefore=3, spaceAfter=6)


def banner(title, subtitle):
    data = [[
        Paragraph(f"<b>{title}</b>", ParagraphStyle(
            "b1", fontName=FONT_REG, fontSize=12.5, leading=16, textColor=white)),
        Paragraph(subtitle, ParagraphStyle(
            "b2", fontName=FONT_REG, fontSize=8.5, leading=11.5, textColor=HexColor("#D7EEE8"))),
    ]]
    t = Table(data, colWidths=[48 * mm, 124 * mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), C_ACCENT),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return t


def tip_box(title, text):
    data = [
        [Paragraph(f"<b>備選／提醒｜{title}</b>", S["tip"])],
        [Paragraph(text, S["tip"])],
    ]
    t = Table(data, colWidths=[172 * mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), C_TIP),
        ("BOX", (0, 0), (-1, -1), 0.5, HexColor("#E0A060")),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (0, 0), 6),
        ("BOTTOMPADDING", (0, -1), (0, -1), 6),
        ("TOPPADDING", (0, 1), (0, 1), 1),
    ]))
    return t


def budget(text):
    return Paragraph(f"<b>當日預算（兩人）</b>：{text}", S["meta"])


def footer(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(C_SOFT)
    canvas.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)
    canvas.setFillColor(C_ACCENT)
    canvas.rect(0, PAGE_H - 7.5 * mm, PAGE_W, 7.5 * mm, fill=1, stroke=0)
    canvas.setFillColor(white)
    canvas.setFont(FONT_REG, 7)
    canvas.drawString(MARGIN, PAGE_H - 5 * mm, "首爾七天休閒度假攻略｜2026.8.8–8.14｜Hotel POCO Seongsu")
    canvas.setFillColor(C_MUTED)
    canvas.setFont(FONT_REG, 7)
    canvas.drawCentredString(PAGE_W / 2, 7.5 * mm, f"— {doc.page} —")
    canvas.setStrokeColor(C_LINE)
    canvas.setLineWidth(0.4)
    canvas.line(MARGIN, 11 * mm, PAGE_W - MARGIN, 11 * mm)
    canvas.restoreState()


def cover(story):
    story.append(Spacer(1, 4 * mm))
    story.append(img("cover-seoul-summer.jpg", 178 * mm, 88 * mm))
    story.append(Spacer(1, 4 * mm))
    ban = Table(
        [[Paragraph("SEOUL · 聖水慢活七天", ParagraphStyle(
            "c1", fontName=FONT_REG, fontSize=20, leading=26, textColor=white, alignment=TA_CENTER))],
         [Paragraph("2026年8月8日（六）— 8月14日（五）｜兩位成人｜休閒度假節奏", ParagraphStyle(
             "c2", fontName=FONT_REG, fontSize=9.5, leading=13, textColor=HexColor("#D7EEE8"), alignment=TA_CENTER))],
         [Paragraph("住宿：Hotel POCO Seongsu｜地鐵2號線聖水站3號出口步行約1分鐘", ParagraphStyle(
             "c3", fontName=FONT_REG, fontSize=8.5, leading=12, textColor=HexColor("#BFE3DA"), alignment=TA_CENTER))]],
        colWidths=[178 * mm],
    )
    ban.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), C_ACCENT),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(ban)
    story.append(Spacer(1, 5 * mm))
    story.append(Paragraph("怎麼用這份攻略", S["h2"]))
    story.append(Paragraph(
        "每天只排 2–3 個重點，其餘留給散步與咖啡。交通以地鐵＋1 公里內步行為主。"
        "8 月首爾盛夏：戶外防曬補水；室內冷氣強，薄外套放包包。中午偏室內，傍晚再戶外。",
        S["body"],
    ))
    story.append(Paragraph("固定不可更動", S["h2"]))
    story.append(Paragraph("• <b>8/9（日）18:00</b>　Alla Prima 生日晚餐（已訂位，套餐約 19–31 萬韓元／人）", S["bullet"]))
    story.append(Paragraph("• <b>8/12（三）</b>　MORE ON HAIR 江南店美髮＋유후네일 신논현美甲", S["bullet"]))
    story.append(Paragraph(
        "航班（以電子機票為準）：去程 <b>UO630</b> HKG→ICN 約 09:45–14:25；"
        "回程 <b>UO631</b> ICN→HKG 約 15:25–16:10。",
        S["body"],
    ))
    story.append(PageBreak())


def overview(story):
    story.append(Paragraph("七天總覽", S["h1"]))
    story.append(hr())
    headers = [Paragraph(f"<b>{x}</b>", S["table_h"]) for x in ("日期", "主題", "重點地點")]
    rows = [
        ["8/8 六 Day1", "抵達聖水・落地慢活", "ICN→Hotel POCO／大林倉庫／성수연방"],
        ["8/9 日 Day2", "麵包午餐・慶生之夜", "首爾林／Cafe Onion／Alla Prima 18:00"],
        ["8/10 一 Day3", "韓牛＋DDP＋一隻雞", "인생한우／DDP／진옥화晚餐"],
        ["8/11 二 Day4", "設計家具＋咖啡日", "안목／Grey Penguin／29CM／LCDC／Object"],
        ["8/12 三 Day5", "美髮＋美甲日", "유후네일／MORE ON HAIR 江南店"],
        ["8/13 四 Day6", "弘大慢逛日", "弘大商圈／延南森林路／炸雞"],
        ["8/14 五 Day7", "退房返港", "周邊早餐／12:00前出發／UO631"],
    ]
    data = [headers] + [[Paragraph(c, S["table"]) for c in r] for r in rows]
    t = Table(data, colWidths=[36 * mm, 46 * mm, 90 * mm])
    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), C_ACCENT),
        ("GRID", (0, 0), (-1, -1), 0.35, C_LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]
    for i in (1, 3, 5, 7):
        style_cmds.append(("BACKGROUND", (0, i), (-1, i), HexColor("#EEF6F3")))
    t.setStyle(TableStyle(style_cmds))
    story.append(t)
    story.append(Spacer(1, 4 * mm))
    story.append(Paragraph("盛夏實用提醒", S["h2"]))
    story.append(Paragraph("• 白天高溫潮濕，SPF50+、帽子、摺疊傘、隨時補水；戶外避開 12:00–15:00 暴晒。", S["bullet"]))
    story.append(Paragraph("• 咖啡廳／地鐵／商場冷氣強，薄外套必備。", S["bullet"]))
    story.append(Paragraph("• 交通卡 T-money；必備 App：Naver Map、Papago、CatchTable、Kakao T。", S["bullet"]))
    story.append(Paragraph("住宿基地", S["h2"]))
    story.append(Paragraph(
        "<b>Hotel POCO Seongsu</b>｜서울 성동구 성수이로 96｜"
        "地鐵2號線 <b>聖水站3號出口</b>右轉約50m｜電話 02-3677-6676／前台 02-462-9610",
        S["body"],
    ))
    story.append(PageBreak())


def day1(story):
    story.append(banner("Day 1｜8月8日（六）", "主題：抵達聖水，落地慢活——只做 check-in、周邊散步、輕鬆晚餐"))
    story.append(Spacer(1, 3 * mm))
    story.append(img("day1-daelim-cafe.jpg"))
    story.append(Paragraph("畫面重點：紅磚倉庫高挑天花＋麵包櫃——飯店同路「성수이로」第一站必拍。", S["caption"]))

    story.append(Paragraph("時間軸", S["h2"]))
    story.append(Paragraph(
        "<b>09:45–14:25｜UO630</b>　香港 HKG → 仁川 ICN（以電子機票為準）。入境＋領行李約 45–75 分。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>約 15:30–17:00｜機場→飯店</b>　AREX 普通車 → <b>弘大入口</b> 轉 2號線 → <b>聖水站</b>3號出口步行約1分。"
        "全程約 70–90 分。預估下午 4–5 點前後抵達。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>17:00–18:00｜Check-in</b>　Hotel POCO Seongsu。沖澡、調冷氣、下載離線地圖。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>18:00–19:30｜飯店周邊散步（同路성수이로）</b><br/>"
        "① <b>대림창고（大林倉庫）</b> 성수이로 78｜紅磚倉庫咖啡／展覽感，聖水站3號出口步行約4–5分，距飯店極近<br/>"
        "② <b>Musinsa Store 성수@대림창고</b> 성수이로 74｜室內冷氣、球鞋牆、服飾（約 11:00–22:00）<br/>"
        "③ <b>성수연방</b> 성수이로14길 14｜複合生活空間；可上3樓 <b>천상가옥</b> 喝一杯（約 11:00–22:00）",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>19:30–21:00｜輕鬆晚餐二選一</b><br/>"
        "A. <b>꿉당 성수</b> 성수이로20길 10｜炭火豬頸肉（約 16:30 起，週末可能排隊）<br/>"
        "B. 飯店步行5分內韓式家常／義大利簡餐｜人均約 1.5–2.5 萬",
        S["body"],
    ))
    story.append(Paragraph("<b>晚上</b>　早點睡，為明天慶生養精神。", S["body"]))
    story.append(budget("餐飲 4–8萬 ＋ 機場交通約 1–1.6萬 ≈ <b>6–10萬韓元</b>"))
    story.append(Spacer(1, 2 * mm))
    story.append(tip_box("延誤／太累", "晚於18:00抵達就外帶回房。大林倉庫人多改坐성수연방吹冷氣。"))
    story.append(PageBreak())


def day2(story):
    story.append(banner("Day 2｜8月9日（日）", "主題：午前樹蔭・Onion麵包午餐・Alla Prima 慶生晚餐"))
    story.append(Spacer(1, 3 * mm))
    story.append(img("day2-seoul-forest.jpg", 178 * mm, 58 * mm))
    story.append(Paragraph("畫面重點：首爾林樹蔭木棧——午前涼爽短走，避開正午暴晒。", S["caption"]))

    story.append(Paragraph("時間軸", S["h2"]))
    story.append(Paragraph(
        "<b>09:30–11:00｜首爾林（서울숲）</b>　只走樹蔭步道 45–60 分。聖水站步行或盆唐線首爾林站。",
        S["body"],
    ))
    story.append(img("day2-onion-bakery.jpg", 178 * mm, 58 * mm))
    story.append(Paragraph("畫面重點：工廠感空間＋麵包山——先拍整櫃再動手；中庭／二樓光線最好。", S["caption"]))

    story.append(Paragraph("<b>11:30–13:30｜午餐：Cafe Onion 성수（網紅麵包店）</b>", S["body"]))
    story.append(Paragraph(
        "• 地址：서울 성동구 아차산로9길 8｜聖水站 <b>2號出口</b> 步行約 2–3 分<br/>"
        "• 營業：平日約 08:00–22:00、週末約 09:00–22:00（L.O. 21:30）<br/>"
        "• 必點：<b>설산 팡도르、앙버터</b>＋冰美式／簽名奶茶｜人均約 1.5–2.5 萬<br/>"
        "• 技巧：兩人分散<b>先佔座再排隊買麵包</b>（週日 13:30 後更擠）<br/>"
        "• 備援：大排長龍就外帶回飯店吃；或改 <b>밀도 성수</b>（왕십리로 96，吐司名店，偏外帶）",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>13:30–16:00｜回飯店</b>　午睡、保濕、化妝穿搭。Alla Prima 建議 smart casual，勿短褲拖鞋。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>16:30 出發｜赴論峴／鶴洞</b>　聖水→建大入口轉7號線→<b>鶴洞站6號出口</b>（約30–40分）。"
        "17:15–17:50 附近慢逛補妝；可選 <b>더스퀘어</b>（鶴洞站步行約4分）坐一下。",
        S["body"],
    ))
    story.append(img("day2-alla-prima-dinner.jpg", 178 * mm, 55 * mm))
    story.append(Paragraph("畫面重點：燭光雙人桌＋精緻擺盤——慶生夜拍菜與碰杯即可，注意店家拍攝禮儀。", S["caption"]))

    story.append(Paragraph("重點：Alla Prima（알라프리마）18:00", S["h2"]))
    story.append(Paragraph(
        "• 地址：서울 강남구 학동로17길 13｜電話 02-511-2555<br/>"
        "• 套餐約 <b>19–31萬韓元／人</b>；兩人含飲品預留 <b>45–75萬</b><br/>"
        "• 晚餐約 2.5–3 小時；遲到約15分可能 no-show<br/>"
        "• 可事先確認生日備註／甜點加字",
        S["body"],
    ))
    story.append(Paragraph("<b>回程</b>　7號線→建大入口→2號線聖水。", S["body"]))
    story.append(budget("Onion 4–6萬 ＋ Alla Prima 45–75萬 ＋ 地鐵 ≈ <b>50–82萬韓元</b>"))
    story.append(Spacer(1, 2 * mm))
    story.append(tip_box("雨天", "取消首爾林，Onion 長坐或성수연방。下午務必留白——今天唯一任務是慶生晚餐。"))
    story.append(PageBreak())


def day3(story):
    story.append(banner("Day 3｜8月10日（一）", "主題：馬場韓牛 → DDP 避暑 → 一隻雞晚餐"))
    story.append(Spacer(1, 3 * mm))
    story.append(img("day3-majang-hanwoo.jpg", 178 * mm, 58 * mm))
    story.append(Paragraph("畫面重點：炭火上的霜降韓牛——拍特寫油脂反光最有「馬場感」。", S["caption"]))

    story.append(Paragraph("時間軸", S["h2"]))
    story.append(Paragraph(
        "<b>11:00 出發｜馬場洞</b>　聖水→往十里轉5號線→<b>馬場站</b>（約20–25分）。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>11:30–14:00｜午餐：인생한우</b><br/>"
        "• 地址：서울 성동구 마장로31길 43 2層｜近馬場畜產市場<br/>"
        "• 怎麼吃：看肉櫃 <b>1++</b>，現場選肉＋上桌費（홀約4,000）再烤<br/>"
        "• 必點：꽃등심、살치살、채끝；收尾물냉면或된장찌개<br/>"
        "• 兩人先點約 600–800g 再加｜預算約 <b>12–25萬</b>｜週末可能候位",
        S["body"],
    ))
    story.append(img("day3-ddp.jpg", 178 * mm, 58 * mm))
    story.append(Paragraph("畫面重點：DDP 銀色曲線外牆——廣角拍建築轉角；室內吹冷氣看當期展。", S["caption"]))

    story.append(Paragraph(
        "<b>14:30–17:30｜東大門設計廣場 DDP</b><br/>"
        "• 交通：馬場／往十里轉乘 → <b>東大門歷史文化公園站</b>（約15–20分）<br/>"
        "• 內容：札哈・哈蒂曲線建築＋當期展覽（付費展查 ddp.or.kr）；公共空間可免費逛<br/>"
        "• 盛夏定位：韓牛後的「室內避暑＋設計感」，不爬山不暴晒<br/>"
        "• 旁邊可顺看 Doota／現代 City Outlet，不強制逛街",
        S["body"],
    ))
    story.append(img("day6-dakhanmari.jpg", 178 * mm, 55 * mm))
    story.append(Paragraph("畫面重點：大鍋滾湯整雞——蒸汽升起瞬間最有儀式感（晚餐檔）。", S["caption"]))

    story.append(Paragraph(
        "<b>18:00–20:00｜晚餐：진옥화할매원조닭한마리（一隻雞）</b><br/>"
        "• 地址：서울 종로구 종로40가길 18｜東大門站 <b>9號出口</b> 步行約6–8分<br/>"
        "• <b>不預約</b>，現場排隊；雙人一隻雞約3萬上下＋年糕／粉條<br/>"
        "• 營業約 10:30–01:00｜DDP 出來走路／短程地鐵即到，動線順",
        S["body"],
    ))
    story.append(Paragraph("<b>回程</b>　地鐵回聖水休息（明天設計家具日）。", S["body"]))
    story.append(budget("韓牛 12–25萬 ＋ DDP／咖啡 1–4萬 ＋ 一隻雞 4–6萬 ≈ <b>17–35萬韓元</b>"))
    story.append(Spacer(1, 2 * mm))
    story.append(tip_box(
        "排隊／雨天",
        "DDP 人多可改隔壁 Doota／現代 City Mall 吹冷氣。"
        "一隻雞排隊過長：同巷其他雞湯店，或回聖水 남다른본가닭한마리（연무장7길 12 2層）。",
    ))
    story.append(PageBreak())


def day4(story):
    story.append(banner("Day 4｜8月11日（二）", "主題：安木湯飯＋聖水設計家具深度遊＋咖啡串連"))
    story.append(Spacer(1, 3 * mm))
    story.append(img("day4-anmok-gukbap.jpg", 178 * mm, 55 * mm))
    story.append(Paragraph("畫面重點：熱湯飯蒸汽＋冷白切肉拼盤——先拍整桌再動筷。", S["caption"]))

    story.append(Paragraph(
        "為什麼週二？Object、아모레성수 週一公休；週二可一次串完。一天認真坐 <b>2 間咖啡</b>即可。",
        S["meta"],
    ))

    story.append(Paragraph("時間軸", S["h2"]))
    story.append(Paragraph(
        "<b>11:30–13:00｜안목 성수</b><br/>"
        "• 주소：뚝섬로13길 34｜聖水3號出口步行約7–10分<br/>"
        "• 雙人：豬國飯×2（各約13,000）＋冰冷모듬수육（약35,000）≈人均3–4萬<br/>"
        "• CatchTable／現場候位；約10:00–22:00（21:00 L.O.）",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>13:30–14:30｜Cafe ① Grey Penguin Company</b><br/>"
        "• 주소：서울숲4길 26-14｜近首爾林｜塔派名店（香草／季節水果塔）＋冰飲<br/>"
        "• 飯後甜點休息、吹冷氣再出發逛家具",
        S["body"],
    ))
    story.append(img("day4-design-shop.jpg", 178 * mm, 55 * mm))
    story.append(Paragraph("畫面重點：大窗採光＋原木家具層次——拍商品陳列比拍人臉更有選品質感。", S["caption"]))

    story.append(Paragraph("<b>14:30–17:30｜設計家具動線</b>", S["body"]))
    story.append(Paragraph(
        "① <b>29CM HOME</b> 연무장길 57｜家具、燈具、生活器物（聖水站步行約5分）<br/>"
        "② <b>29CM HOME 2</b> 연무장길 110｜食品雜貨／廚房浴室情境展區<br/>"
        "③ <b>LCDC SEOUL</b> 연무장17길 10｜複合文化＋選品／快閃<br/>"
        "④ <b>Object</b> 서울숲길 36 2層｜約12:30–20:30，獨立設計文具小物",
        S["bullet"],
    ))
    story.append(Paragraph(
        "<b>中段 Cafe ②（彈性插入）</b>　逛累就進 <b>대림창고</b> 성수이로 78 咖啡，"
        "或 <b>성수연방・천상가옥</b> 성수이로14길 14 3層（空間大好坐）。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>可選</b>　<b>아모레성수</b> 아차산로11길 7｜10:30–20:30（週一休→今天可）試香、看窗邊綠意。",
        S["body"],
    ))
    story.append(Paragraph("<b>晚上</b>　輕食回飯店；明天美容日，少油炸。", S["body"]))
    story.append(budget("안목 6–8萬 ＋ 咖啡甜點 4–7萬 ＋ 小物 0–15萬 ≈ <b>10–30萬韓元</b>"))
    story.append(Spacer(1, 2 * mm))
    story.append(tip_box("店休／下雨", "縮成兩間設計店＋長坐咖啡（Grey Penguin＋천상가옥）。大雨不硬走연무장全線。"))
    story.append(PageBreak())


def day5(story):
    story.append(banner("Day 5｜8月12日（三）", "主題：美髮＋美甲——新論峴一站搞定"))
    story.append(Spacer(1, 3 * mm))
    story.append(img("day5-beauty-salon.jpg"))
    story.append(Paragraph("畫面重點：完成後拍雙手凝膠＋新髮型側臉最實用。", S["caption"]))

    story.append(Paragraph(
        "MORE ON HAIR <b>江南店（강남대로 475）</b>≠江南總店，預約請認準地址。"
        "週一、週二公休——今天（週三）可服務。服務時段以 Naver 預約確認單為準。",
        S["body"],
    ))

    story.append(Paragraph("時間軸（先美甲→再美髮）", S["h2"]))
    story.append(Paragraph(
        "<b>10:00–10:40｜移動</b>　聖水2號線直達 <b>新論峴</b>（約20–25分），2號出口出站。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>11:00–13:00｜美甲：유후네일 신논현</b><br/>"
        "• 주소：서울 강남구 봉은사로1길 37｜新論峴步行約5–8分<br/>"
        "• Naver Map 搜尋：<b>유후네일 신논현</b> → 開啟「預約／N예약」<br/>"
        "• 與美髮店同商圈，步行銜接｜凝膠約4–7萬起",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>13:00–14:00｜輕午餐</b>　江南大路商圈簡餐，避免重口味沾手。人均1.2–2萬。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>14:00–18:00｜美髮：MORE ON HAIR 江南店</b><br/>"
        "• 주소：서울 서초구 강남대로 475, 3F<br/>"
        "• 新論峴站 <b>2號出口</b> 步行可達｜出示預約、確認剪染燙與價位",
        S["body"],
    ))
    story.append(Paragraph("<b>晚上</b>　回聖水輕食／蛋糕慶祝新造型。", S["body"]))
    story.append(budget("美甲4–10萬 ＋ 美髮（剪染燙常見10–40萬＋）＋ 餐3–5萬"))
    story.append(Spacer(1, 2 * mm))
    story.append(tip_box(
        "美甲約滿",
        "Naver「신논현역 네일샵」篩預約可；或미쥬네일（봉은사로4길 23一帶）。"
        "優先保住 MORE ON HAIR 時段。",
    ))
    story.append(PageBreak())


def day6(story):
    story.append(banner("Day 6｜8月13日（四）", "主題：弘大慢逛日——商圈＋延南森林路＋炸雞夜"))
    story.append(Spacer(1, 3 * mm))
    story.append(img("day6-hongdae.jpg", 178 * mm, 60 * mm))
    story.append(Paragraph("畫面重點：弘大街景樹蔭與店面——午後多進店躲暑，傍晚再走森林路。", S["caption"]))

    story.append(Paragraph(
        "一隻雞已安排在 Day 3 晚餐。今天專心拍弘大／延南，節奏放慢。",
        S["meta"],
    ))

    story.append(Paragraph("時間軸", S["h2"]))
    story.append(Paragraph(
        "<b>10:30–11:30｜出發</b>　聖水站 <b>2號線直達弘大入口</b>（약30분）。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>11:30–13:00｜早午餐＋咖啡</b>　弘大入口站出站後選一間商圈咖啡／早午餐，"
        "避開最擠巷口正中央，往側街較好坐。人均1.5–2.5萬。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>13:00–17:30｜弘大→延南洞慢逛</b><br/>"
        "① <b>弘大商圈</b>：服飾、彩妝、選品店——盛夏策略是「多進店、少站街」<br/>"
        "② 往北步行至 <b>延南洞（연남동）</b><br/>"
        "③ <b>京義線森林路延南段（경의선숲길）</b>：樹蔭步道散步＋找咖啡坐45–60分（比主街舒適好拍）",
        S["body"],
    ))
    story.append(img("day6-fried-chicken.jpg", 178 * mm, 55 * mm))
    story.append(Paragraph("畫面重點：脆皮炸雞＋冰啤酒——弘大夜食收尾；暖光金屬盤最好拍。", S["caption"]))

    story.append(Paragraph(
        "<b>18:30–20:30｜炸雞夜（留弘大，少折返）</b><br/>"
        "• 建議：Naver Map 搜「<b>홍대 치킨</b>」依當日評分／距離選店，配生啤<br/>"
        "• 預算：兩人約4–7萬<br/>"
        "• 若執意回聖水：레츠잇치킨（왕십리로4길 12）或 교촌치킨 성수역점",
        S["body"],
    ))
    story.append(Paragraph("<b>晚上</b>　回聖水收拾明日托運行李，確認 UO631 航廈。", S["body"]))
    story.append(budget("餐飲咖啡 5–8萬 ＋ 炸雞啤酒 4–7萬 ＋ 地鐵／小物 2–8萬 ≈ <b>11–23萬韓元</b>"))
    story.append(Spacer(1, 2 * mm))
    story.append(tip_box(
        "暴晒／雷雨",
        "主街太曬就縮短逛街，改延南洞咖啡長坐＋森林路傍晚再走。"
        "大雨改室內選品店／生活商場，炸雞外帶回飯店亦可。",
    ))
    story.append(PageBreak())


def day7(story):
    story.append(banner("Day 7｜8月14日（五）", "主題：退房返港——中午前出發，不排完整景點"))
    story.append(Spacer(1, 3 * mm))
    story.append(img("day7-icn-airport.jpg"))
    story.append(Paragraph("畫面重點：仁川機場明亮大廳——留足時間比再塞景點重要。", S["caption"]))

    story.append(Paragraph("時間軸", S["h2"]))
    story.append(Paragraph(
        "<b>08:30–10:00｜收拾＋早餐</b>　飯店周邊：Onion 外帶／편의店／천상가옥輕食。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>10:00–11:30｜最後採購（可選）</b>　僅限步行10分內藥妝或大林倉庫／Musinsa；不排完整景點。"
        "液體與刀剪注意托運規定。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>建議 12:00–12:30前退房出發</b><br/>"
        "回程 <b>UO631</b>：ICN→HKG，起飛約 <b>15:25–16:10</b>（以機票為準）。<br/>"
        "路線：聖水→弘大入口轉 AREX→仁川（約70–90分）＋提前2小時候機。",
        S["body"],
    ))
    story.append(Paragraph("<b>機場</b>　報到、托運、退稅；航廈內可當最後一餐。", S["body"]))
    story.append(budget("早餐2–4萬 ＋ 機場交通約1–1.6萬 ≈ <b>3–9萬韓元</b>"))
    story.append(Spacer(1, 2 * mm))
    story.append(tip_box("突發", "AREX異常改機場巴士或計程車。起飛前3小時再確認航廈與航班狀態。"))
    story.append(PageBreak())


def prep(story):
    story.append(Paragraph("行前準備清單", S["h1"]))
    story.append(hr())
    story.append(Paragraph("天氣與穿著（8月首爾盛夏）", S["h2"]))
    story.append(Paragraph(
        "白天常見高溫約30–35°C、濕度高，偶有雷陣雨。透氣衣物、涼鞋、帽、防晒；包包放薄外套。"
        "中午偏室內（DDP、商場、選品店），傍晚再戶外。",
        S["body"],
    ))
    story.append(Paragraph("金錢與匯率", S["h2"]))
    story.append(Paragraph(
        "以行前銀行／Google匯率為準（常見約1 HKD≈170–185 KRW，會波動）。"
        "信用卡普及；備5–10萬韓元現金。本攻略預算為兩人當日粗估。",
        S["body"],
    ))
    story.append(Paragraph("上網", S["h2"]))
    story.append(Paragraph("機場領 eSIM／旅遊卡（5–7天）。Hotspot 可只開一張給兩人用。", S["body"]))
    story.append(Paragraph("必備 App", S["h2"]))
    story.append(Paragraph("• <b>Naver Map</b>：找路、營業時間、美甲美髮預約", S["bullet"]))
    story.append(Paragraph("• <b>Papago</b>：韓中翻譯／菜單拍照", S["bullet"]))
    story.append(Paragraph("• <b>CatchTable</b>：안목等餐廳候位", S["bullet"]))
    story.append(Paragraph("• <b>Kakao T</b>：偶發叫車｜航空公司 App 查航班", S["bullet"]))
    story.append(Paragraph("出發前48小時再確認", S["h2"]))
    story.append(Paragraph("• Alla Prima 18:00 訂位姓名、服裝、過敏原", S["bullet"]))
    story.append(Paragraph("• MORE ON HAIR 江南店（강남대로475）預約單", S["bullet"]))
    story.append(Paragraph("• 유후네일 Naver 預約", S["bullet"]))
    story.append(Paragraph("• UO630／UO631 航廈與起飛時間", S["bullet"]))
    story.append(Paragraph("• 飯店 check-in／out 與行李寄放", S["bullet"]))
    story.append(Spacer(1, 5 * mm))
    end = Table(
        [[Paragraph(
            "祝兩位在聖水過一個不趕路的夏天——把最好的精神留給 8/9 慶生晚餐，其餘日子慢慢走就好。",
            ParagraphStyle("end", fontName=FONT_REG, fontSize=9.5, leading=14,
                           textColor=white, alignment=TA_CENTER)
        )]],
        colWidths=[172 * mm],
    )
    end.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), C_ACCENT),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
    ]))
    story.append(end)
    story.append(Spacer(1, 3 * mm))
    story.append(Paragraph(
        "免責：店家營業時間、價格、公休日與航班時刻可能變動；請以 Naver Map／官方預約／電子機票為準。",
        S["caption"],
    ))


def main():
    doc = SimpleDocTemplate(
        str(OUT), pagesize=A4,
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=12 * mm, bottomMargin=14 * mm,
        title="首爾七天遊圖文攻略 2026.8.8–8.14",
        author="Seoul Leisure Guide",
    )
    story = []
    cover(story)
    overview(story)
    day1(story)
    day2(story)
    day3(story)
    day4(story)
    day5(story)
    day6(story)
    day7(story)
    prep(story)
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    main()
