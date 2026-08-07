#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generate Traditional Chinese Seoul 7-day travel guide PDF."""

from pathlib import Path
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white, black
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle,
    PageBreak, KeepTogether, HRFlowable, ListFlowable, ListItem,
)
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

ROOT = Path(__file__).resolve().parent
IMG = ROOT / "images_opt"
OUT = ROOT / "artifacts" / "首爾七天遊圖文攻略_2026.pdf"
OUT.parent.mkdir(parents=True, exist_ok=True)

# Prefer local CJK fonts
FONT_REG = "NotoSansCJK"
FONT_BOLD = "NotoSansCJK"
for path in (
    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
    "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf",
):
    p = Path(path)
    if p.exists():
        try:
            pdfmetrics.registerFont(TTFont("GuideFont", str(p), subfontIndex=0))
            FONT_REG = "GuideFont"
            FONT_BOLD = "GuideFont"
            break
        except Exception:
            continue
else:
    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    FONT_REG = "STSong-Light"
    FONT_BOLD = "STSong-Light"

# Colors — summer Seoul editorial (teal / warm sand / charcoal; avoid purple/cream clichés)
C_BG = HexColor("#F7F3EE")
C_INK = HexColor("#1C2421")
C_MUTED = HexColor("#5A655F")
C_ACCENT = HexColor("#0E6B5C")
C_ACCENT2 = HexColor("#C45C26")
C_SOFT = HexColor("#E6EFEA")
C_LINE = HexColor("#D5D0C8")
C_CARD = HexColor("#FFFFFF")
C_TIP = HexColor("#FFF6E8")

PAGE_W, PAGE_H = A4
MARGIN = 16 * mm


def styles():
    s = {}
    s["cover_brand"] = ParagraphStyle(
        "cover_brand", fontName=FONT_BOLD, fontSize=28, leading=34,
        textColor=white, alignment=TA_CENTER, spaceAfter=6,
    )
    s["cover_sub"] = ParagraphStyle(
        "cover_sub", fontName=FONT_REG, fontSize=12, leading=18,
        textColor=HexColor("#E8F5F1"), alignment=TA_CENTER, spaceAfter=4,
    )
    s["h1"] = ParagraphStyle(
        "h1", fontName=FONT_BOLD, fontSize=18, leading=24,
        textColor=C_ACCENT, spaceBefore=4, spaceAfter=8,
    )
    s["h2"] = ParagraphStyle(
        "h2", fontName=FONT_BOLD, fontSize=13, leading=18,
        textColor=C_INK, spaceBefore=10, spaceAfter=4,
    )
    s["h3"] = ParagraphStyle(
        "h3", fontName=FONT_BOLD, fontSize=11, leading=15,
        textColor=C_ACCENT2, spaceBefore=6, spaceAfter=3,
    )
    s["body"] = ParagraphStyle(
        "body", fontName=FONT_REG, fontSize=9.5, leading=14.5,
        textColor=C_INK, alignment=TA_JUSTIFY, spaceAfter=4,
    )
    s["meta"] = ParagraphStyle(
        "meta", fontName=FONT_REG, fontSize=9, leading=13,
        textColor=C_MUTED, spaceAfter=3,
    )
    s["caption"] = ParagraphStyle(
        "caption", fontName=FONT_REG, fontSize=8, leading=11,
        textColor=C_MUTED, alignment=TA_CENTER, spaceBefore=2, spaceAfter=8,
    )
    s["bullet"] = ParagraphStyle(
        "bullet", fontName=FONT_REG, fontSize=9.5, leading=14,
        textColor=C_INK, leftIndent=8, spaceAfter=2,
    )
    s["tip"] = ParagraphStyle(
        "tip", fontName=FONT_REG, fontSize=9, leading=13,
        textColor=C_INK, spaceAfter=2,
    )
    s["table"] = ParagraphStyle(
        "table", fontName=FONT_REG, fontSize=8, leading=11,
        textColor=C_INK, alignment=TA_LEFT,
    )
    s["table_h"] = ParagraphStyle(
        "table_h", fontName=FONT_BOLD, fontSize=8, leading=11,
        textColor=white, alignment=TA_CENTER,
    )
    s["footer"] = ParagraphStyle(
        "footer", fontName=FONT_REG, fontSize=7.5, leading=10,
        textColor=C_MUTED, alignment=TA_CENTER,
    )
    return s


S = styles()


def img(name, w=170 * mm, h=78 * mm):
    stem = Path(name).stem
    path = IMG / f"{stem}.jpg"
    if not path.exists():
        path = IMG / name
    if not path.exists():
        return Spacer(1, 2)
    return Image(str(path), width=w, height=h, kind="proportional")


def hr():
    return HRFlowable(width="100%", thickness=0.6, color=C_LINE, spaceBefore=4, spaceAfter=8)


def section_banner(title, subtitle):
    data = [[Paragraph(f"<b>{title}</b>", ParagraphStyle(
        "ban", fontName=FONT_BOLD, fontSize=14, leading=18, textColor=white
    )), Paragraph(subtitle, ParagraphStyle(
        "ban2", fontName=FONT_REG, fontSize=9, leading=12, textColor=HexColor("#D7EEE8"), alignment=TA_LEFT
    ))]]
    t = Table(data, colWidths=[55 * mm, 120 * mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), C_ACCENT),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return t


def tip_box(title, text):
    data = [
        [Paragraph(f"<b>備選／提醒｜{title}</b>", S["tip"])],
        [Paragraph(text, S["tip"])],
    ]
    t = Table(data, colWidths=[175 * mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), C_TIP),
        ("BOX", (0, 0), (-1, -1), 0.5, HexColor("#E0A060")),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (0, 0), 7),
        ("BOTTOMPADDING", (0, -1), (0, -1), 7),
        ("TOPPADDING", (0, 1), (0, 1), 2),
    ]))
    return t


def budget_line(text):
    return Paragraph(f"<b>當日預算（兩人）</b>：{text}", S["meta"])


def add_header_footer(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(C_SOFT)
    canvas.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)
    canvas.setFillColor(C_ACCENT)
    canvas.rect(0, PAGE_H - 8 * mm, PAGE_W, 8 * mm, fill=1, stroke=0)
    canvas.setFillColor(white)
    canvas.setFont(FONT_REG, 7.5)
    canvas.drawString(MARGIN, PAGE_H - 5.2 * mm, "首爾七天休閒度假攻略｜2026.8.8–8.14｜Hotel POCO Seongsu")
    canvas.setFillColor(C_MUTED)
    canvas.setFont(FONT_REG, 7.5)
    canvas.drawCentredString(PAGE_W / 2, 8 * mm, f"— {doc.page} —")
    canvas.setStrokeColor(C_LINE)
    canvas.setLineWidth(0.4)
    canvas.line(MARGIN, 12 * mm, PAGE_W - MARGIN, 12 * mm)
    canvas.restoreState()


def cover_page(story):
    # Full-bleed-ish cover using stacked elements
    story.append(Spacer(1, 8 * mm))
    story.append(img("cover-seoul-summer.png", 178 * mm, 95 * mm))
    story.append(Spacer(1, 6 * mm))
    banner = Table(
        [[Paragraph("SEOUL · 聖水慢活七天", ParagraphStyle(
            "cb", fontName=FONT_BOLD, fontSize=22, leading=28, textColor=white, alignment=TA_CENTER
        ))],
         [Paragraph("2026年8月8日（六）— 8月14日（五）｜兩位成人｜休閒度假節奏", ParagraphStyle(
             "cs", fontName=FONT_REG, fontSize=10, leading=14, textColor=HexColor("#D7EEE8"), alignment=TA_CENTER
         ))],
         [Paragraph("住宿：Hotel POCO Seongsu｜地鐵2號線聖水站3號出口步行約1分鐘", ParagraphStyle(
             "cs2", fontName=FONT_REG, fontSize=9, leading=13, textColor=HexColor("#BFE3DA"), alignment=TA_CENTER
         ))]],
        colWidths=[178 * mm],
    )
    banner.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), C_ACCENT),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
    ]))
    story.append(banner)
    story.append(Spacer(1, 8 * mm))
    story.append(Paragraph("這份攻略給誰用？", S["h2"]))
    story.append(Paragraph(
        "兩位成人、以休閒度假為主：每天只排 2–3 個重點，其餘留給散步、咖啡與臨時起意。"
        "交通以地鐵＋1 公里內步行為主，盡量少搭計程車。8 月首爾是盛夏，戶外移動請防曬補水；"
        "室內冷氣強，薄外套或開衫很實用。",
        S["body"],
    ))
    story.append(Paragraph("固定不可更動行程", S["h2"]))
    story.append(Paragraph("• <b>8/9（日）18:00</b>　Alla Prima 生日晚餐（已訂位，套餐約 19–31 萬韓元／人）", S["bullet"]))
    story.append(Paragraph("• <b>8/12（三）</b>　MORE ON HAIR 江南店美髮＋新論峴周邊美甲（銜接安排）", S["bullet"]))
    story.append(Spacer(1, 4 * mm))
    story.append(Paragraph(
        "航班（實際時刻以電子機票為準）：去程 <b>UO630</b> HKG→ICN 約 09:45–14:25；"
        "回程 <b>UO631</b> ICN→HKG 約 15:25–16:10（起飛窗口可能微調）。",
        S["body"],
    ))
    story.append(PageBreak())


def overview_page(story):
    story.append(Paragraph("七天總覽（先看這頁再出發）", S["h1"]))
    story.append(hr())
    headers = [
        Paragraph("<b>日期</b>", S["table_h"]),
        Paragraph("<b>主題</b>", S["table_h"]),
        Paragraph("<b>重點地點</b>", S["table_h"]),
    ]
    rows = [
        ["8/8 六 Day1", "抵達聖水・慢活安頓", "ICN→Hotel POCO／聖水散步／輕鬆晚餐"],
        ["8/9 日 Day2", "輕鬆午前・慶生之夜", "首爾林／咖啡／Alla Prima（18:00）"],
        ["8/10 一 Day3", "馬場洞韓牛體驗", "馬場韓牛街／聖水咖啡收尾"],
        ["8/11 二 Day4", "聖水設計家具深度遊", "Object／29CM HOME／LCDC／安木湯飯"],
        ["8/12 三 Day5", "美髮＋美甲日", "MORE ON HAIR 江南店／유후네일"],
        ["8/13 四 Day6", "一隻雞＋炸雞夜", "珍玉花奶奶一隻雞／聖水炸雞"],
        ["8/14 五 Day7", "退房返港", "周邊早餐／12:00前出發／UO631"],
    ]
    data = [headers]
    for r in rows:
        data.append([Paragraph(x, S["table"]) for x in r])
    t = Table(data, colWidths=[38 * mm, 48 * mm, 89 * mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), C_ACCENT),
        ("BACKGROUND", (0, 1), (-1, 1), HexColor("#EEF6F3")),
        ("BACKGROUND", (0, 3), (-1, 3), HexColor("#EEF6F3")),
        ("BACKGROUND", (0, 5), (-1, 5), HexColor("#EEF6F3")),
        ("BACKGROUND", (0, 7), (-1, 7), HexColor("#EEF6F3")),
        ("GRID", (0, 0), (-1, -1), 0.4, C_LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(t)
    story.append(Spacer(1, 6 * mm))
    story.append(Paragraph("盛夏實用提醒（每天都適用）", S["h2"]))
    story.append(Paragraph("• 白天高溫潮濕，建議 SPF50+、帽子、摺疊傘、隨時補水；戶外活動避開 12:00–15:00 暴晒時段。", S["bullet"]))
    story.append(Paragraph("• 咖啡廳／地鐵／商場冷氣強，薄外套或開衫放包包；出汗後可回飯店沖澡再出發。", S["bullet"]))
    story.append(Paragraph("• 交通卡：T-money 或開卡後綁 Naver Pay／信用卡；單程地鐵約 1,500–2,000 韓元起。", S["bullet"]))
    story.append(Paragraph("• 必備 App：Naver Map（找路／預約）、KakaoMap（備援）、Papago（翻譯）、CatchTable（餐廳候位）。", S["bullet"]))
    story.append(Spacer(1, 4 * mm))
    story.append(Paragraph("住宿基地", S["h2"]))
    story.append(Paragraph(
        "<b>Hotel POCO Seongsu（호텔 포코 성수）</b><br/>"
        "地址：서울 성동구 성수이로 96（Seongsu-ro 96, Seongdong-gu）<br/>"
        "交通：地鐵2號線 <b>聖水站（성수）3號出口</b> 右轉步行約 50m（約1分鐘）。"
        "電話：02-3677-6676／前台 02-462-9610。",
        S["body"],
    ))
    story.append(PageBreak())


def day1(story):
    story.append(section_banner("Day 1｜8月8日（六）", "主題：抵達聖水，慢活安頓——今天只做「落地＋吃飯＋散步」"))
    story.append(Spacer(1, 4 * mm))
    story.append(img("day1-seongsu-street.png"))
    story.append(Paragraph("畫面重點：聖水紅磚倉庫街＋夏日樹蔭，飯店周邊最適合落地後的第一段慢走。", S["caption"]))

    story.append(Paragraph("時間軸", S["h2"]))
    story.append(Paragraph(
        "<b>09:45–14:25｜航班 UO630</b>　香港 HKG → 仁川 ICN（實際時刻以電子機票確認信為準）。"
        "落地後預留入境、領行李約 45–75 分鐘。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>約 15:30–17:00｜機場 → 飯店</b><br/>"
        "推薦：<b>AREX 機場鐵道（普通）</b> 仁川機場 → <b>弘大入口（홍대입구）</b> 轉 <b>2號線內圈／往聖水方向</b> → "
        "<b>聖水站</b> 3號出口步行約1分鐘。<br/>"
        "全程約 70–90 分鐘。亦可搭機場巴士後轉地鐵，但行李多時 AREX＋地鐵較可控。"
        "抵達時間預估 <b>下午4–5點</b> 前後。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>17:00–18:00｜Check-in</b>　Hotel POCO Seongsu。放下行李、沖澡、調冷氣，確認 Naver Map 離線常用地點。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>18:00–19:30｜飯店周邊散步</b>　3號出口周邊紅磚巷、選品小店、咖啡外帶。路程控制在 1 公里內來回，"
        "不排景點打卡清單。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>19:30–21:00｜輕鬆晚餐</b>　建議飯店步行可達的韓式家常／義大利簡餐／拉麵任選一間；"
        "人均約 1.5–2.5 萬韓元。吃完買瓶水與隔日早餐零食回房。",
        S["body"],
    ))
    story.append(Paragraph("<b>晚上</b>　早點睡，為明天慶生晚餐養精神。", S["body"]))
    story.append(budget_line("餐飲 4–6萬 ＋ 機場交通約 1–1.6萬 ＋ 雜費 1萬 ≈ <b>6–9萬韓元</b>"))
    story.append(Spacer(1, 3 * mm))
    story.append(tip_box(
        "航班延誤／體力不夠",
        "若抵達晚於18:00，改成便利店／飯店附近外帶回房；明天才開始正式逛街。"
        "若AREX人潮多，可改機場巴士6013→建大入口再轉2號線聖水（行李多時較省力但時間未必較短）。",
    ))
    story.append(PageBreak())


def day2(story):
    story.append(section_banner("Day 2｜8月9日（日）", "主題：輕鬆午前・慶生之夜——下午留白，為 Alla Prima 18:00 蓄勢"))
    story.append(Spacer(1, 4 * mm))
    story.append(img("day2-seoul-forest.png", 178 * mm, 72 * mm))
    story.append(Paragraph("畫面重點：首爾林樹蔭木棧道，午前涼爽散步的最佳取景——側光葉影＋手持冰飲。", S["caption"]))

    story.append(Paragraph("時間軸", S["h2"]))
    story.append(Paragraph(
        "<b>10:00–12:00｜首爾林（서울숲）輕走</b><br/>"
        "地鐵：聖水站步行可至公園北側入口，或搭 <b>盆唐線／水仁盆唐線 首爾林站</b>。"
        "只走樹蔭步道 45–60 分鐘，拍照後離開，不排滿園區。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>12:00–13:30｜午餐＋咖啡</b>　首爾林／聖水咖啡館輕食即可（沙拉、三明治、冷麵擇一）。"
        "人均 1.5–2.5 萬。下午要化妝穿搭，午餐勿吃太撐。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>13:30–16:00｜回飯店休息</b>　午睡／保濕／燙髮定妝／挑衣服。"
        "Alla Prima 建議 smart casual；避免短褲拖鞋運動服。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>16:30 出發｜前往論峴／鶴洞</b><br/>"
        "路線：聖水站 <b>2號線</b> → 建大入口轉 <b>7號線</b> → <b>鶴洞站（학동）6號出口</b> 步行約5–8分鐘；"
        "或 2號線 → 教大轉 3號線 → 狎鷗亭，再步行／短程接駁。全程約 30–40 分鐘。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>17:15–17:50｜餐廳附近慢逛</b>　提前抵達論峴家具街周邊，找咖啡坐一下、補妝、確認訂位姓名。"
        "勿踩點卡到 18:00——遲到超過約15分鐘可能被視為 no-show。",
        S["body"],
    ))

    story.append(img("day2-alla-prima-dinner.png", 178 * mm, 72 * mm))
    story.append(Paragraph(
        "畫面重點：燭光、精緻擺盤、雙人餐桌——慶生夜拍「菜＋手／碰杯」比拍整間店更有氣氛（注意店家拍攝禮儀）。",
        S["caption"],
    ))

    story.append(Paragraph("重點：Alla Prima（알라프리마）生日晚餐", S["h2"]))
    story.append(Paragraph(
        "• 地址：서울 강남구 학동로17길 13（Hakdong-ro 17-gil 13, Gangnam-gu）<br/>"
        "• 時間：<b>18:00</b> 入座（已訂位）；晚餐約需 2.5–3 小時<br/>"
        "• 預算：套餐約 <b>19–31萬韓元／人</b>（視當日菜單／酒款）；兩人含服務與飲品可預留 <b>45–75萬</b><br/>"
        "• 氣氛小提醒：訂位時可再確認是否備註生日，詢問甜點／卡片加字服務；手機調靜音，享受一餐節奏。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>21:00–22:00｜返程</b>　7號線鶴洞 → 建大入口轉 2號線 → 聖水。回飯店可買氣泡水慶祝收工。",
        S["body"],
    ))
    story.append(budget_line("午餐咖啡 4–6萬 ＋ Alla Prima 45–75萬 ＋ 地鐵約 0.5萬 ≈ <b>50–82萬韓元</b>"))
    story.append(Spacer(1, 3 * mm))
    story.append(tip_box(
        "雨天／體力不足",
        "午前改成聖水室內選品店或飯店附近咖啡坐到中午；取消首爾林。Alla Prima 不可改期時，"
        "下午務必留白——今天唯一任務是把慶生晚餐吃漂亮。",
    ))
    story.append(PageBreak())


def day3(story):
    story.append(section_banner("Day 3｜8月10日（一）", "主題：馬場洞韓牛街——一餐吃懂選肉與現場炭烤"))
    story.append(Spacer(1, 4 * mm))
    story.append(img("day3-majang-hanwoo.png"))
    story.append(Paragraph(
        "畫面重點：桌爐炭火、霜降韓牛油脂反光、煙霧與金屬夾——拍特寫最有「馬場感」。",
        S["caption"],
    ))

    story.append(Paragraph("時間軸", S["h2"]))
    story.append(Paragraph(
        "<b>10:30–11:30｜慢起＋咖啡</b>　聖水站周邊外帶咖啡，補水防暑。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>11:30 出發｜前往馬場洞</b><br/>"
        "推薦路線：聖水站 <b>2號線</b> → <b>往十里（왕십리）</b> 轉 <b>5號線</b> → <b>馬場站（마장）1號出口</b>；"
        "步行進入韓牛／肉鋪巷約 5–10 分鐘。全程約 20–25 分鐘。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>12:00–14:30｜馬場洞韓牛午餐（避開晚餐尖峰）</b><br/>"
        "怎麼選店：看門口肉櫃是否標示 <b>1++／1+</b>、現場可指定部位、店內有排煙良好的烤桌。"
        "可在 Naver Map 搜尋「마장동 한우」依即時評分＋「本地人多」篩選；"
        "熱門如馬場畜產相關烤肉店、한우 정육식당類型皆可。<br/>"
        "必點部位：<b>燈芯（등심）</b>、<b>腹脇／附骨（갈비살／본갈비）</b>、<b>雪花上腦（살치살）</b>；"
        "可加一份冷麵或湯收尾。<br/>"
        "預算：兩人約 <b>12–25萬韓元</b>（依等級與份量）；先點 600–800g 再加點，避免浪費。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>15:00–17:30｜返聖水・咖啡漫步</b>　原路回聖水，找有冷氣的咖啡廳坐 1–2 小時（推薦沿阿且山路／首爾林路一帶）。"
        "今天不排第二個重餐。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>晚上</b>　若還餓，便利店三角飯糰／水果即可；為明天家具日養腳力。",
        S["body"],
    ))
    story.append(budget_line("韓牛 12–25萬 ＋ 咖啡點心 3–5萬 ＋ 地鐵約 0.5萬 ≈ <b>16–31萬韓元</b>"))
    story.append(Spacer(1, 3 * mm))
    story.append(tip_box(
        "排隊過長／想吹冷氣",
        "改往十里站周邊大型商場（如往十里站複合商場）用餐後再決定是否前往馬場；"
        "或改訂聖水／建大預約制韓牛店。下雨則縮短戶外巷弄停留，直攻有冷氣的定食烤肉店。",
    ))
    story.append(PageBreak())


def day4(story):
    story.append(section_banner("Day 4｜8月11日（二）", "主題：聖水設計家具／生活選品深度遊＋安木豬肉湯飯"))
    story.append(Spacer(1, 4 * mm))
    story.append(img("day4-design-shop.png", 178 * mm, 70 * mm))
    story.append(Paragraph(
        "畫面重點：大面積採光窗＋原木家具陳列——拍「商品層次＋窗光」比拍人臉更有選品店質感。",
        S["caption"],
    ))

    story.append(Paragraph("為什麼排在週二？", S["h3"]))
    story.append(Paragraph(
        "Object 聖水店每週一公休；週二開始逛可一次串起家具／生活選品動線，避免走回頭路。",
        S["body"],
    ))

    story.append(Paragraph("時間軸", S["h2"]))
    story.append(Paragraph(
        "<b>10:30–11:30｜咖啡啟動</b>　飯店出發，沿聖水站往首爾林／煙霧長路方向慢走。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>11:30–13:00｜午餐：安木 돼지국밥／안목 성수</b><br/>"
        "地址：서울 성동구 뚝섬로13길 34（聖水站3號出口步行約7–10分鐘）<br/>"
        "特色：釜山風格豬肉湯飯，湯頭濃但乾淨；搭配「冷白切肉／모듬 수육」最完整。<br/>"
        "雙人建議：豬國飯×2（各約 13,000）＋冰冷拼盤수육（약 35,000）≈ <b>人均 3–4萬</b>。"
        "也可點熱수육。午餐尖峰可能需候位，可用 CatchTable／現場遙控候位。<br/>"
        "營業參考：約 10:00–22:00（21:00 last order，以現場為準）。",
        S["body"],
    ))
    story.append(img("day4-anmok-gukbap.png", 178 * mm, 68 * mm))
    story.append(Paragraph("畫面重點：熱湯飯蒸汽＋冷白切肉拼盤對比——先拍整桌再動筷。", S["caption"]))

    story.append(Paragraph("<b>13:30–17:30｜設計家具／選品路線（可彈性縮短）</b>", S["body"]))
    story.append(Paragraph(
        "1) <b>29CM HOME Seongsu（이구홈 성수）</b><br/>"
        "　주소：서울 성동구 연무장길 57｜聖水站步行約5分｜家具、燈具、生活器物實品試摸的主力店。<br/>"
        "　同品牌延伸：연무장길 110 的 <b>29CM HOME 2</b>（食品雜貨／廚房浴室情境展區）。",
        S["bullet"],
    ))
    story.append(Paragraph(
        "2) <b>LCDC SEOUL</b><br/>"
        "　주소：서울 성동구 연무장17길 10｜複合文化空間＋選品／快閃，適合吹冷氣、看展、買小物。",
        S["bullet"],
    ))
    story.append(Paragraph(
        "3) <b>Object（오브젝트）聖水店</b><br/>"
        "　주소：서울 성동구 서울숲길 36 2층｜營業約 12:30–20:30，<b>週一公休</b>｜"
        "獨立設計文具與生活小物，DIY 組合有趣，適合帶回手禮。",
        S["bullet"],
    ))
    story.append(Paragraph(
        "動線建議：煙霧長路（연무장길）一帶逛完 → 往首爾林路方向到 Object，全程步行多在 1 公里級，可穿插咖啡。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>晚上</b>　回飯店休息；若想宵夜，明天美髮美甲前避免太油的炸物，留到 Day 6。",
        S["body"],
    ))
    story.append(budget_line("安木 6–8萬 ＋ 咖啡 2–4萬 ＋ 選品／小物 0–15萬 ＋ 地鐵少 ≈ <b>8–27萬韓元</b>"))
    story.append(Spacer(1, 3 * mm))
    story.append(tip_box(
        "店休日／下雨",
        "若某店臨時公休，改走聖水「서울숲길／연무장」其他選品店或 MUJI／大型生活館吹冷氣。"
        "大雨天把路線縮成兩間店＋長坐咖啡，家具改看型錄不硬走。",
    ))
    story.append(PageBreak())


def day5(story):
    story.append(section_banner("Day 5｜8月12日（三）", "主題：美髮＋美甲日——新論峴站一站搞定，少搬運"))
    story.append(Spacer(1, 4 * mm))
    story.append(img("day5-beauty-salon.png"))
    story.append(Paragraph(
        "畫面重點：明亮鏡面、凝膠美甲特寫、整潔工位——完成後拍雙手＋新髮型側臉最實用。",
        S["caption"],
    ))

    story.append(Paragraph("預約原則", S["h2"]))
    story.append(Paragraph(
        "MORE ON HAIR 江南店（강남대로 475）與江南總店是不同分店，預約請認準地址。"
        "營業參考 10:00–19:00；服務時段可能約至 13:00–22:00（以 Naver 預約確認單為準）。"
        "週一、週二公休——今天（週三）可服務。",
        S["body"],
    ))

    story.append(Paragraph("時間軸（建議：先美甲 → 再美髮）", S["h2"]))
    story.append(Paragraph(
        "<b>10:00–10:40｜移動</b><br/>"
        "聖水站 <b>2號線</b> 直達 <b>新論峴站（신논현）</b>，約 20–25 分鐘；"
        "<b>2號出口</b>出站，步行至美甲店。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>11:00–13:00｜美甲（建議提前在 Naver 預約）</b><br/>"
        "推薦：<b>유후네일 신논현점（YouWho Nail Sinnonhyeon）</b><br/>"
        "地址：서울 강남구 봉은사로1길 37（論峴洞）｜距新論峴站步行約 5–8 分鐘，"
        "與 MORE ON HAIR（강남대로 475）同屬 2號出口商圈，可步行銜接。<br/>"
        "Naver Map 搜尋關鍵字：<b>유후네일 신논현</b> 或 <b>봉은사로1길 37 네일</b>；"
        "在 App 內開啟「預約／N예약」選時段。<br/>"
        "預算參考：手部凝膠／今月藝術約 4–7萬起（依款式）；手＋足套組更高。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>13:00–14:00｜輕午餐</b>　新論峴／江南大路商圈簡餐（避免重口味沾手）。人均 1.2–2萬。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>14:00–18:00｜美髮：MORE ON HAIR 江南店</b><br/>"
        "地址：서울 서초구 강남대로 475, 3F（475 Gangnam-daero, Seocho-gu）<br/>"
        "交通：新論峴站 <b>2號出口</b> 步行可達。<br/>"
        "到店出示預約、確認染燙／剪髮內容與價位；過程可補水休息。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>晚上</b>　回聖水輕食或蛋糕慶祝新造型；早點休息。",
        S["body"],
    ))
    story.append(budget_line("美甲 4–10萬 ＋ 美髮（依項目，常見剪染燙 10–40萬＋）＋ 餐飲 3–5萬 ＋ 地鐵 ≈ 視服務項目"))
    story.append(Spacer(1, 3 * mm))
    story.append(tip_box(
        "美甲店約滿／公休",
        "備選同商圈：Naver Map 搜尋「신논현역 네일샵」並篩選「預約可」；"
        "或「미쥬네일」（봉은사로4길 23 一帶，距新論峴約5–10分）。"
        "若美髮延遲，美甲改排美髮前一晚或回程前一天——優先保住 MORE ON HAIR 時段。",
    ))
    story.append(PageBreak())


def day6(story):
    story.append(section_banner("Day 6｜8月13日（四）", "主題：東大門一隻雞名店＋聖水炸雞夜——最後完整玩一天"))
    story.append(Spacer(1, 4 * mm))
    story.append(img("day6-dakhanmari.png", 178 * mm, 70 * mm))
    story.append(Paragraph(
        "畫面重點：大鍋滾湯、整雞與粉條年糕——蒸汽升起來的瞬間最有「一隻雞」儀式感。",
        S["caption"],
    ))

    story.append(Paragraph("時間軸", S["h2"]))
    story.append(Paragraph(
        "<b>10:30–11:30｜出發東大門商圈</b><br/>"
        "聖水站 <b>2號線</b> → <b>東大門歷史文化公園站</b> 或 <b>東大門站</b>；"
        "改路線：2號線 → 往十里轉其他線亦可。預留步行找店時間。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>11:30–13:30｜一隻雞：珍玉花奶奶元祖一隻雞</b><br/>"
        "店名：진옥화할매원조닭한마리（Jin Ok-hwa Halmae Dakhanmari）<br/>"
        "地址：서울 종로구 종로40가길 18｜近東大門綜合市場美食巷<br/>"
        "交通：<b>東大門站 9號出口</b> 步行約 6–8 分鐘；或鍾路5街站附近步行。<br/>"
        "提醒：<b>不接受預約</b>，現場抽號／排隊；週末與正餐時間常見候位。"
        "建議開門前後到，或錯開 12:00 尖峰。<br/>"
        "點法：雙人點 <b>一隻雞</b>（約 3萬上下，以現場為準）＋年糕／粉條／烏龍麵加料；"
        "最後湯底可加麵收尾。營業約 10:30–01:00。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>14:00–17:00｜附近慢逛或回聖水休息</b>　東大門 Design Plaza（DDP）室內吹冷氣看建築曲線；"
        "或直接回飯店午睡，為晚上炸雞留胃。",
        S["body"],
    ))
    story.append(img("day6-fried-chicken.png", 178 * mm, 68 * mm))
    story.append(Paragraph("畫面重點：脆皮炸雞＋冰啤酒金屬盤——夜晚室內暖光最有「聖水夜食」感。", S["caption"]))

    story.append(Paragraph(
        "<b>18:30–20:30｜炸雞：聖水 Let's Eat Chicken（레츠잇치킨）</b><br/>"
        "地址：서울 성동구 왕십리로4길 12｜<b>纛島站（뚝섬）7號出口</b> 步行約 3–5 分鐘<br/>"
        "招牌：美式風格炸雞／LEC 炸雞、Buffalo、拼盤；適合配生啤。<br/>"
        "預算：兩人約 4–7萬（含飲料）。<br/>"
        "備援（更韓式經典）：<b>橋村炸雞 聖水站店（교촌치킨 성수역점）</b>　"
        "아차산로7길 10-1｜招牌：蜂蜜綜合／醬油半半。",
        S["body"],
    ))
    story.append(Paragraph("<b>晚上</b>　回飯店收拾明日要托運的行李，確認回程航班與機場交通。", S["body"]))
    story.append(budget_line("一隻雞 4–6萬 ＋ 炸雞啤酒 4–7萬 ＋ 咖啡雜費 2–4萬 ＋ 地鐵 ≈ <b>10–17萬韓元</b>"))
    story.append(Spacer(1, 3 * mm))
    story.append(tip_box(
        "珍玉花排隊過長",
        "同巷還有多家一隻雞可選；或改回聖水「남다른본가닭한마리 성수점」"
        "（연무장7길 12 2층，聖水站4號出口步行約2–3分）吃湯雞，把東大門改成 DDP 看展半日。",
    ))
    story.append(PageBreak())


def day7(story):
    story.append(section_banner("Day 7｜8月14日（五）", "主題：退房返港——上午只做早餐／最後採購，中午前出發"))
    story.append(Spacer(1, 4 * mm))
    story.append(img("day7-icn-airport.png"))
    story.append(Paragraph("畫面重點：仁川機場明亮大廳與行李——留足時間比再塞一個景點重要。", S["caption"]))

    story.append(Paragraph("時間軸", S["h2"]))
    story.append(Paragraph(
        "<b>08:30–10:00｜收拾行李＋飯店周邊早餐</b>　咖啡＋三明治／粥品即可；可到便利店補韓國餅乾、海苔作為伴手禮。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>10:00–11:30｜最後採購（可選）</b>　僅限聖水站步行 10 分鐘內的藥妝／選品；"
        "<b>不要排完整景點</b>。液體與刀剪注意托運規定。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>建議 12:00–12:30 前退房出發</b><br/>"
        "回程航班 <b>UO631</b>：首爾 ICN → 香港 HKG，起飛約 <b>15:25–16:10</b>（以機票為準），"
        "抵達香港約 18:15–19:00。<br/>"
        "交通：聖水站 <b>2號線</b> → <b>弘大入口</b> 轉 <b>AREX</b> → 仁川機場第1／第2航廈（依航空公司航廈確認）。"
        "車程約 <b>70–90 分鐘</b>；加上提前 <b>2 小時</b> 抵達機場，中午前出發最穩妥。",
        S["body"],
    ))
    story.append(Paragraph(
        "<b>13:30–15:00｜機場</b>　報到、托運、安檢、換匯／退稅（若有）。"
        "航廈內用餐當最後一餐亦可。",
        S["body"],
    ))
    story.append(budget_line("早餐 2–4萬 ＋ 機場交通約 1–1.6萬 ＋ 機場小食 0–3萬 ≈ <b>3–9萬韓元</b>"))
    story.append(Spacer(1, 3 * mm))
    story.append(tip_box(
        "出發前突發狀況",
        "若 AREX 異常，改機場巴士或計程車（費用明顯較高，兩人行李多時可考慮）。"
        "起飛前 3 小時再確認航廈與航班狀態（HK Express App／機場看板）。",
    ))
    story.append(PageBreak())


def prep_page(story):
    story.append(Paragraph("行前準備清單", S["h1"]))
    story.append(hr())

    story.append(Paragraph("天氣與穿著（8月首爾盛夏）", S["h2"]))
    story.append(Paragraph(
        "白天常見高溫約 30–35°C、濕度高，偶有午後雷陣雨。穿著：透氣棉麻、涼鞋、帽、防晒；"
        "包包放薄外套應付室內冷氣。每日步行不多，但地鐵出站仍有日照路段。",
        S["body"],
    ))

    story.append(Paragraph("金錢與匯率", S["h2"]))
    story.append(Paragraph(
        "以行前銀行／Google 匯率為準（常見約 1 HKD ≈ 170–185 KRW 區間，會波動）。"
        "信用卡在首爾普及；小攤與部分美甲店可能偏好轉帳／現金，備 5–10萬韓元現金較安心。"
        "本攻略預算皆為「兩人當日」粗估。",
        S["body"],
    ))

    story.append(Paragraph("上網：SIM／eSIM／WiFi", S["h2"]))
    story.append(Paragraph(
        "推薦機場領 eSIM／實體旅遊卡（5–7天無限流量）。Hotspot 分享給同伴可只開一張。"
        "飯店有 WiFi，但出門导航仍需行動數據。",
        S["body"],
    ))

    story.append(Paragraph("必備 App", S["h2"]))
    story.append(Paragraph("• <b>Naver Map</b>：找路、營業時間、美甲／美髮預約（N예약）", S["bullet"]))
    story.append(Paragraph("• <b>Papago</b>：韓中翻譯（菜單拍照翻譯很好用）", S["bullet"]))
    story.append(Paragraph("• <b>CatchTable</b>：熱門餐廳候位／訂位（安木等）", S["bullet"]))
    story.append(Paragraph("• <b>Kakao T</b>：偶發需要叫車時", S["bullet"]))
    story.append(Paragraph("• <b>航空公司 App／HK Express</b>：值機與航班異動", S["bullet"]))

    story.append(Paragraph("出發前 48 小時再確認", S["h2"]))
    story.append(Paragraph("• Alla Prima 訂位姓名、時間 18:00、服裝與過敏原", S["bullet"]))
    story.append(Paragraph("• MORE ON HAIR 江南店（강남대로 475）預約確認單", S["bullet"]))
    story.append(Paragraph("• 유후네일（또는 備選美甲）Naver 預約", S["bullet"]))
    story.append(Paragraph("• UO630／UO631 航廈與起飛時間", S["bullet"]))
    story.append(Paragraph("• 飯店 check-in／check-out 時間與行李寄放", S["bullet"]))

    story.append(Spacer(1, 6 * mm))
    story.append(Paragraph("美甲預約快速指引（可截圖）", S["h2"]))
    story.append(Paragraph(
        "主選：유후네일 신논현점｜서울 강남구 봉은사로1길 37<br/>"
        "Naver Map 搜尋：유후네일 신논현<br/>"
        "銜接：新論峴站2號出口商圈 ↔ MORE ON HAIR（강남대로 475, 3F）步行可達。<br/>"
        "建議時序：11:00 美甲 → 午餐 → 14:00 美髮（可依預約單微調對調）。",
        S["body"],
    ))

    story.append(Spacer(1, 8 * mm))
    end = Table(
        [[Paragraph(
            "祝兩位在聖水過一個不趕路的夏天——把最好的精神留給 8/9 慶生晚餐，其餘日子慢慢走就好。",
            ParagraphStyle("end", fontName=FONT_REG, fontSize=10, leading=15,
                           textColor=white, alignment=TA_CENTER)
        )]],
        colWidths=[178 * mm],
    )
    end.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), C_ACCENT),
        ("TOPPADDING", (0, 0), (-1, -1), 12),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 12),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("RIGHTPADDING", (0, 0), (-1, -1), 12),
    ]))
    story.append(end)
    story.append(Spacer(1, 4 * mm))
    story.append(Paragraph(
        "免責：店家營業時間、價格、公休日與航班時刻可能變動；出行前請以 Naver Map／官方預約／電子機票為準。",
        S["caption"],
    ))


def main():
    doc = SimpleDocTemplate(
        str(OUT),
        pagesize=A4,
        leftMargin=MARGIN,
        rightMargin=MARGIN,
        topMargin=14 * mm,
        bottomMargin=16 * mm,
        title="首爾七天遊圖文攻略 2026.8.8–8.14",
        author="Seoul Leisure Guide",
    )
    story = []
    cover_page(story)
    overview_page(story)
    day1(story)
    day2(story)
    day3(story)
    day4(story)
    day5(story)
    day6(story)
    day7(story)
    prep_page(story)
    doc.build(story, onFirstPage=add_header_footer, onLaterPages=add_header_footer)
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    main()
