#!/usr/bin/env python3
"""Generate a leisure Sakura trip PDF: Kanazawa & Toyama, Apr 6–10."""

from pathlib import Path

from fpdf import FPDF

FONT = "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"
OUT = Path(__file__).resolve().parent / "金澤富山賞櫻之旅_4月6-10日.pdf"

# Soft sakura palette (avoid purple/cream AI defaults)
INK = (45, 42, 40)
MUTED = (110, 100, 98)
SAKURA = (196, 112, 128)
SAKURA_SOFT = (245, 232, 234)
LEAF = (90, 122, 98)
CREAM_BG = (252, 248, 246)
LINE = (220, 205, 200)
WHITE = (255, 255, 255)
CARD_BG = (255, 252, 250)


class TripPDF(FPDF):
    def __init__(self):
        super().__init__(format="A4", unit="mm")
        self.set_auto_page_break(auto=True, margin=18)
        self.add_font("zh", "", FONT)
        self.add_font("zh", "B", FONT)

    def header(self):
        if self.page_no() == 1:
            return
        self.set_font("zh", "", 9)
        self.set_text_color(*MUTED)
        self.set_y(10)
        self.cell(0, 6, "金澤 · 富山｜休閒賞櫻之旅  4/6–4/10", align="L")
        self.set_draw_color(*SAKURA)
        self.set_line_width(0.3)
        self.line(12, 17, 198, 17)
        self.ln(10)

    def footer(self):
        self.set_y(-14)
        self.set_font("zh", "", 8)
        self.set_text_color(*MUTED)
        self.cell(0, 8, f"{self.page_no()}", align="C")

    def section_title(self, text):
        self.set_font("zh", "B", 15)
        self.set_text_color(*SAKURA)
        self.cell(0, 10, text, new_x="LMARGIN", new_y="NEXT")
        self.set_draw_color(*SAKURA)
        self.set_line_width(0.4)
        y = self.get_y()
        self.line(12, y, 48, y)
        self.ln(5)

    def body(self, text, size=10.5, leading=6.2):
        self.set_font("zh", "", size)
        self.set_text_color(*INK)
        self.multi_cell(0, leading, text)
        self.ln(1.5)

    def muted(self, text, size=9.5, leading=5.5):
        self.set_font("zh", "", size)
        self.set_text_color(*MUTED)
        self.multi_cell(0, leading, text)
        self.ln(1)

    def bullet(self, title, detail=""):
        x = self.get_x()
        y = self.get_y()
        self.set_fill_color(*SAKURA)
        self.ellipse(x + 1.2, y + 2.0, 2.2, 2.2, style="F")
        self.set_xy(x + 7, y)
        self.set_font("zh", "B", 10.5)
        self.set_text_color(*INK)
        if detail:
            self.cell(0, 6.2, title, new_x="LMARGIN", new_y="NEXT")
            self.set_x(x + 7)
            self.set_font("zh", "", 9.5)
            self.set_text_color(*MUTED)
            self.multi_cell(0, 5.4, detail)
        else:
            self.multi_cell(0, 6.2, title)
        self.ln(1.2)

    def day_banner(self, day_label, title, subtitle):
        if self.get_y() > 240:
            self.add_page()
        self.ln(2)
        x, y = 12, self.get_y()
        w, h = 186, 22
        self.set_fill_color(*SAKURA_SOFT)
        self.rect(x, y, w, h, style="F")
        self.set_fill_color(*SAKURA)
        self.rect(x, y, 3.2, h, style="F")
        self.set_xy(x + 8, y + 3.5)
        self.set_font("zh", "B", 11)
        self.set_text_color(*SAKURA)
        self.cell(0, 6, day_label)
        self.set_xy(x + 8, y + 10.5)
        self.set_font("zh", "B", 13)
        self.set_text_color(*INK)
        self.cell(120, 7, title)
        self.set_font("zh", "", 9.5)
        self.set_text_color(*MUTED)
        self.cell(0, 7, subtitle, align="R")
        self.set_y(y + h + 4)

    def time_row(self, time, place, tip=""):
        self.set_font("zh", "B", 10)
        self.set_text_color(*LEAF)
        self.cell(22, 6.2, time)
        self.set_font("zh", "B", 10.5)
        self.set_text_color(*INK)
        self.multi_cell(0, 6.2, place)
        if tip:
            self.set_x(self.l_margin + 22)
            self.set_font("zh", "", 9.2)
            self.set_text_color(*MUTED)
            self.multi_cell(0, 5.2, tip)
        self.ln(1.8)

    def tip_box(self, text):
        self.ln(1)
        y = self.get_y()
        self.set_font("zh", "", 9.5)
        # measure
        self.set_xy(16, y + 3)
        # draw after measuring approximate height
        lines = max(2, len(text) // 42 + text.count("\n") + 1)
        h = 6 + lines * 5.2
        self.set_fill_color(238, 244, 238)
        self.rect(12, y, 186, h, style="F")
        self.set_xy(16, y + 3)
        self.set_text_color(*LEAF)
        self.multi_cell(178, 5.2, "小提醒｜" + text)
        self.set_y(y + h + 3)


def build():
    pdf = TripPDF()
    pdf.set_margins(12, 16, 12)

    # ── Cover ──
    pdf.add_page()
    pdf.set_fill_color(*CREAM_BG)
    pdf.rect(0, 0, 210, 297, style="F")

    # top wash
    pdf.set_fill_color(*SAKURA_SOFT)
    pdf.rect(0, 0, 210, 118, style="F")
    pdf.set_fill_color(*SAKURA)
    pdf.rect(0, 118, 210, 2.5, style="F")

    pdf.set_y(36)
    pdf.set_font("zh", "", 12)
    pdf.set_text_color(*SAKURA)
    pdf.cell(0, 8, "HOKURIKU  SAKURA  JOURNEY", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)
    pdf.set_font("zh", "B", 28)
    pdf.set_text_color(*INK)
    pdf.cell(0, 14, "金澤 · 富山", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("zh", "B", 22)
    pdf.set_text_color(*SAKURA)
    pdf.cell(0, 12, "休閒賞櫻之旅", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)
    pdf.set_font("zh", "", 13)
    pdf.set_text_color(*MUTED)
    pdf.cell(0, 8, "4 月 6 日 — 4 月 10 日｜五天四夜", align="C", new_x="LMARGIN", new_y="NEXT")

    pdf.set_y(140)
    pdf.set_font("zh", "", 11)
    pdf.set_text_color(*INK)
    pdf.set_x(28)
    pdf.multi_cell(
        154,
        7,
        "節奏放慢、少趕景點。以兼六園與松川的櫻花為主軸，"
        "穿插茶屋街漫步、溫泉放鬆與富山灣海鮮，讓北陸春天慢慢展開。",
        align="C",
    )

    pdf.set_y(175)
    info = [
        ("行程風格", "休閒賞櫻 · 步行友善 · 彈性安排"),
        ("主要據點", "金澤 3 晚 → 富山 1 晚"),
        ("交通建議", "北陸新幹線／小松機場／富山機場"),
        ("賞櫻重點", "兼六園、金澤城、松川千本櫻"),
    ]
    for label, val in info:
        pdf.set_x(40)
        pdf.set_font("zh", "B", 10.5)
        pdf.set_text_color(*SAKURA)
        pdf.cell(32, 8, label)
        pdf.set_font("zh", "", 10.5)
        pdf.set_text_color(*INK)
        pdf.cell(0, 8, val, new_x="LMARGIN", new_y="NEXT")

    pdf.set_y(230)
    pdf.set_draw_color(*LINE)
    pdf.line(50, 230, 160, 230)
    pdf.set_y(238)
    pdf.set_font("zh", "", 9)
    pdf.set_text_color(*MUTED)
    pdf.cell(0, 6, "櫻花盛開約在 4 月初至中旬，實際開花請出發前查氣象廳／當地情報", align="C")
    pdf.cell(0, 6, "本行程為規劃建議，可依航班與體力微調", align="C", new_x="LMARGIN", new_y="NEXT")

    # ── Overview ──
    pdf.add_page()
    pdf.section_title("行程總覽")
    overview = [
        ("Day 1｜4/6", "抵達金澤", "入住、近江町市場、夜間散步"),
        ("Day 2｜4/7", "金澤賞櫻全日", "兼六園、金澤城、東茶屋街"),
        ("Day 3｜4/8", "金澤慢遊", "長町武家屋敷、21 世紀美術館、溫泉"),
        ("Day 4｜4/9", "前往富山", "松川千本櫻、環水公園、海鮮晚餐"),
        ("Day 5｜4/10", "返程", "早晨散步後離開"),
    ]
    for day, title, desc in overview:
        y = pdf.get_y()
        pdf.set_fill_color(*CARD_BG)
        pdf.set_draw_color(*LINE)
        pdf.rect(12, y, 186, 14, style="DF")
        pdf.set_xy(16, y + 3.5)
        pdf.set_font("zh", "B", 10.5)
        pdf.set_text_color(*SAKURA)
        pdf.cell(42, 7, day)
        pdf.set_text_color(*INK)
        pdf.cell(40, 7, title)
        pdf.set_font("zh", "", 9.5)
        pdf.set_text_color(*MUTED)
        pdf.cell(0, 7, desc)
        pdf.set_y(y + 16)

    pdf.ln(4)
    pdf.section_title("交通與住宿建議")
    pdf.bullet("抵達方式", "台灣出發可飛小松（金澤近）或富山機場；也可經東京／大阪轉北陸新幹線。返程可從富山或小松離開。")
    pdf.bullet("區域交通", "金澤市區巴士（一日券方便）、計程車短程；金澤⇄富山約 20–25 分新幹線，或特急／高速巴士。")
    pdf.bullet("住宿配置", "金澤住 3 晚（市區步行可達兼六園／片町），富山住 1 晚（車站或松川周邊）。節奏較鬆，行李只搬一次。")
    pdf.bullet("預算感（兩人參考）", "住宿約 ¥12,000–22,000／晚；餐食 ¥4,000–8,000／人／日；賞櫻景點多為公園免費或低門票。")

    pdf.ln(2)
    pdf.section_title("行李與季節提醒")
    pdf.bullet("天氣", "北陸 4 月初仍涼，白天約 10–16°C，早晚可低於 8°C。薄羽絨＋圍巾＋防水鞋較安心。")
    pdf.bullet("雨天備案", "金澤 soft rain 常見；美術館、茶屋、溫泉、金箔體驗都適合。")
    pdf.bullet("賞櫻禮儀", "勿折枝、勿進圍籬拍照；熱門時段兼六園與松川較擠，建議晨間或傍晚。")

    # ── Day 1 ──
    pdf.add_page()
    pdf.day_banner("DAY 1　4 月 6 日", "抵達金澤，慢慢落地", "金澤")
    pdf.time_row("午前", "抵達小松機場／金澤站", "機場巴士約 40–50 分到金澤站；新幹線則直達站內。先寄放行李或直接 check-in。")
    pdf.time_row("午後", "近江町市場輕食", "牡蠣、壽司、加賀野菜；吃到飽即可，留胃給晚上。")
    pdf.time_row("傍晚", "香林坊・片町散步", "沿著犀川河岸走走，感受金澤「用水之城」的節奏。")
    pdf.time_row("晚上", "加賀料理或居酒屋", "推薦嚐じぶ煮、加賀蓮藕、のどぐろ。早點休息，為明天賞櫻留體力。")
    pdf.tip_box("今晚若抵達較早，可順路看金澤城公園外圍燈景；不必進園，輕鬆即可。")

    # ── Day 2 ──
    pdf.day_banner("DAY 2　4 月 7 日", "兼六園櫻花全日", "金澤・賞櫻主軸")
    pdf.time_row("07:30", "兼六園晨間入園", "清晨人少、光線柔，是拍照與散步最佳時段。園內「霞ヶ池」「唐崎松」周邊櫻景經典。")
    pdf.time_row("10:00", "金澤城公園", "與兼六園相連，城牆與石川門襯櫻花。可在公園草坪休息，不必趕行程。")
    pdf.time_row("12:30", "午餐（兼六園周邊）", "茶屋甜點（金澤咖啡、和菓子）或金澤咖哩皆可。")
    pdf.time_row("14:30", "東茶屋街", "慢走木造街屋、金箔霜淇淋；可預約金箔貼箔體驗（約 1 小時）。")
    pdf.time_row("17:30", "回飯店休息／自由", "不想趕夜櫻就回飯店；想拍夜櫻可再訪兼六園夜間開園日（季節限定，出發前確認）。")
    pdf.tip_box("兼六園門票約 ¥320；持文化一日券可含城址等多館。雨天改逛博物館與茶屋，櫻花仍可隔窗欣賞。")

    # ── Day 3 ──
    pdf.day_banner("DAY 3　4 月 8 日", "金澤慢遊與放鬆", "金澤")
    pdf.time_row("午前", "長町武家屋敷跡", "土塀、水路與殘雪融後的春色，適合慢拍。附近可順訪奈良茶屋或咖啡。")
    pdf.time_row("午間", "金澤 21 世紀美術館", "建築本身就是景點；展覽隨檔期，逛累了就坐中庭休息。")
    pdf.time_row("午後", "自由選一", "① 大野日和山展望（海風）　② 金箔工房體驗　③ 市區咖啡與選物店")
    pdf.time_row("傍晚", "溫泉放鬆", "片町周邊溫泉設施，或入住飯店泡湯；為明天轉往富山養精神。")
    pdf.time_row("晚上", "輕食即可", "不必大餐；準備隔日退房與行李。")
    pdf.tip_box("這天刻意留白：賞櫻旅行最容易「景點過剩」。午睡、逛街、看書都算行程的一部分。")

    # ── Day 4 ──
    pdf.add_page()
    pdf.day_banner("DAY 4　4 月 9 日", "轉往富山，松川賞櫻", "金澤 → 富山")
    pdf.time_row("09:00", "退房，前往富山", "北陸新幹線かがやき／はくたか約 20–25 分；行李可寄コインロッカー後先賞櫻。")
    pdf.time_row("10:30", "松川公園・千本櫻", "富山代表性櫻花名所，沿松川約 2.5 km 櫻並木。可搭遊覽船（季節運行）從水上仰望。")
    pdf.time_row("12:30", "城址公園周邊午餐", "富山ブラック拉麵或白蝦丼，分量適中即可。")
    pdf.time_row("15:00", "富岩運河環水公園", "星巴克草皮店與運河散步；天氣好時適合坐著發呆看春天。")
    pdf.time_row("晚上", "富山灣海鮮晚餐", "推薦：白蝦、螢火蟲魷（季節）、寒鰤尾聲／春ぶり、寿司カウンター。")
    pdf.tip_box("松川夜間常有燈籠點燈（櫻花季），若體力允許可晚飯後再走一回夜櫻，氣氛很不同。")

    # ── Day 5 ──
    pdf.day_banner("DAY 5　4 月 10 日", "早晨散步，踏上歸途", "富山 → 返程")
    pdf.time_row("午前", "視航班／新幹線彈性安排", "時間充裕：再訪松川晨櫻，或到富山玻璃美術館（運河旁）短暫參觀。")
    pdf.time_row("離開前", "伴手禮採買", "富山：ホワイトショコラ、ますのすし；金澤若尚未買：ひがし茶屋街金箔甜點、加賀棒茶。")
    pdf.time_row("返程", "富山機場／新幹線接續", "富山機場赴台航班視當季；或新幹線至東京／大阪再轉機。預留安檢與轉乘緩衝。")
    pdf.tip_box("若回程經金澤／小松，可把大件伴手禮留到最後一天車站百貨購買，減少一路拖行李。")

    # ── Practical ──
    pdf.ln(3)
    pdf.section_title("賞櫻景點速查")
    spots = [
        ("兼六園（金澤）", "日本三名園之一；晨間最美。門票約 ¥320。"),
        ("金澤城公園", "與兼六園相連；城門與石垣配櫻。多區域免費。"),
        ("松川公園（富山）", "約 450 株櫻並木；可搭遊船。夜間常有點燈。"),
        ("富岩運河環水公園", "運河、草地、建築美照；偏休閒散步。"),
        ("高岡古城公園（選逛）", "若多留半天，可從富山電車約 15–20 分前往。"),
    ]
    for name, desc in spots:
        pdf.bullet(name, desc)

    pdf.ln(2)
    pdf.section_title("美食清單（隨興點）")
    pdf.body(
        "金澤：じぶ煮、のどぐろ定食、金澤咖哩、甜蝦丼、金箔霜淇淋、加賀風和菓子。\n"
        "富山：白蝦、ますのすし、富山ブラック、氷見うどん、新鮮寿司カウンター。"
    )

    pdf.ln(2)
    pdf.section_title("彈性備案")
    pdf.bullet("櫻花未滿開", "仍可走園內松與池景；改加重美術館、茶屋、溫泉比例。")
    pdf.bullet("遇雨", "金澤：21 世紀美術館、鈴木大拙館、茶屋街室內店；富山：玻璃美術館、站內商場。")
    pdf.bullet("想多看海", "Day 3 改半日雨晴海岸或冰見方向（車程較長，僅體力好時建議）。")
    pdf.bullet("想泡溫泉一晚", "可把其中一晚改住金澤近郊溫泉旅館（要犧牲市區便利）。")

    # ── Closing checklist ──
    pdf.add_page()
    pdf.section_title("出發前檢查清單")
    checks = [
        "確認櫻花滿開預測（氣象廳・金澤／富山）",
        "訂妥金澤 3 晚＋富山 1 晚住宿",
        "訂好機票／北陸新幹線（自由席或指定席）",
        "下載 Google Maps／Yahoo!乗換案内 離線可用",
        "準備現金少量（市場小店仍方便）＋交通 IC 卡",
        "薄暖外套、折傘、好走的鞋",
        "飯店／機場接送或巴士時刻再確認一次",
    ]
    for c in checks:
        y = pdf.get_y()
        pdf.set_draw_color(*SAKURA)
        pdf.set_line_width(0.4)
        pdf.rect(14, y + 1.2, 4.2, 4.2)
        pdf.set_xy(22, y)
        pdf.set_font("zh", "", 10.5)
        pdf.set_text_color(*INK)
        pdf.cell(0, 6.5, c, new_x="LMARGIN", new_y="NEXT")
        pdf.ln(1.5)

    pdf.ln(8)
    pdf.set_fill_color(*SAKURA_SOFT)
    pdf.rect(12, pdf.get_y(), 186, 42, style="F")
    y = pdf.get_y()
    pdf.set_xy(20, y + 8)
    pdf.set_font("zh", "B", 13)
    pdf.set_text_color(*SAKURA)
    pdf.cell(0, 8, "旅行一句話", new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(20)
    pdf.set_font("zh", "", 11)
    pdf.set_text_color(*INK)
    pdf.multi_cell(
        170,
        6.5,
        "北陸的春天不吵鬧。把行程留給櫻花、水路與一頓好好吃的飯，就足夠了。",
    )

    pdf.ln(16)
    pdf.set_font("zh", "", 9)
    pdf.set_text_color(*MUTED)
    pdf.cell(0, 6, "本行程手冊僅供規劃參考｜實際開花、營業與交通請以當地最新資訊為準", align="C")

    pdf.output(str(OUT))
    print(f"Wrote {OUT}")
    return OUT


if __name__ == "__main__":
    build()
