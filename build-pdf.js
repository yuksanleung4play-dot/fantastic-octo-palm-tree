/*
 * PDF 產生器
 * 讀取 data.js，排版成 A4 PDF：首爾質感傢俬探店地圖（聖水洞 × 漢南洞）
 *
 * 用法：
 *   npm install puppeteer-core
 *   node build-pdf.js
 * （需系統已安裝 Google Chrome 與 Noto CJK 字型）
 */

const fs = require("fs");
const path = require("path");
const puppeteer = require("puppeteer-core");

const OUT = path.join(__dirname, "首爾質感傢俬探店地圖.pdf");
const CHROME =
  process.env.CHROME_PATH ||
  ["/usr/local/bin/google-chrome", "/usr/bin/google-chrome", "/usr/bin/chromium"].find(
    function (p) { return fs.existsSync(p); }
  );

// 載入 data.js（瀏覽器全域變數 DISTRICTS），不需修改原檔
function loadData() {
  const code = fs.readFileSync(path.join(__dirname, "data.js"), "utf8");
  return new Function(code + "; return DISTRICTS;")();
}

const TIER_LABEL = { 1: "₩", 2: "₩₩", 3: "₩₩₩" };
const esc = function (s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
};

function dist(a, b) {
  return Math.sqrt(Math.pow(a.x - b.x, 2) + Math.pow(a.y - b.y, 2));
}
function walkMinutes(a, b) {
  return Math.max(3, Math.round(dist(a, b) / 55) + 2);
}
function sorted(d) {
  return d.shops.slice().sort(function (a, b) { return a.order - b.order; });
}

function buildMap(d) {
  const shops = sorted(d);
  let s = '<svg viewBox="0 0 960 640" class="map" xmlns="http://www.w3.org/2000/svg">';
  // 底
  s += '<rect x="0" y="0" width="960" height="640" rx="18" fill="#efe7d8"/>';
  // 街道
  d.mapStreets.forEach(function (st) {
    s += '<line x1="' + st.x1 + '" y1="' + st.y1 + '" x2="' + st.x2 + '" y2="' + st.y2 +
      '" stroke="#d8c9b2" stroke-width="14" stroke-linecap="round"/>';
    s += '<text x="' + (st.x1 + st.x2) / 2 + '" y="' + ((st.y1 + st.y2) / 2 - 12) +
      '" fill="#a7967c" font-size="16" font-weight="500">' + esc(st.label) + "</text>";
  });
  // 路線
  s += '<polyline points="' + shops.map(function (x) { return x.x + "," + x.y; }).join(" ") +
    '" fill="none" stroke="#b5552f" stroke-width="3" stroke-dasharray="9 8" opacity="0.75"/>';
  // 標記
  shops.forEach(function (x) {
    const anchor = x.x > 780 ? "end" : "start";
    const nx = x.x > 780 ? x.x - 24 : x.x + 24;
    s += '<circle cx="' + x.x + '" cy="' + x.y + '" r="18" fill="#b5552f"/>';
    s += '<text x="' + x.x + '" y="' + x.y + '" fill="#fff" font-size="18" font-weight="700" text-anchor="middle" dominant-baseline="central">' + x.order + "</text>";
    s += '<text x="' + nx + '" y="' + (x.y + 5) + '" fill="#2b2723" font-size="15" font-weight="500" text-anchor="' + anchor + '">' + esc(x.name) + "</text>";
  });
  s += "</svg>";
  return s;
}

function buildRoute(d) {
  const shops = sorted(d);
  let li = "";
  shops.forEach(function (x, i) {
    const walk = i === 0 ? "由最近車站起步" : "步行約 " + walkMinutes(shops[i - 1], x) + " 分鐘";
    li += '<li><span class="r-num">' + x.order + '</span><span class="r-name">' + esc(x.name) +
      '</span><span class="r-walk">↳ ' + esc(walk) + "</span></li>";
  });
  return '<ol class="route">' + li + "</ol>";
}

function buildCard(x) {
  const tags = x.highlights.map(function (h) { return "<span>" + esc(h) + "</span>"; }).join("");
  return (
    '<article class="card">' +
      '<div class="c-top"><span class="c-num">' + x.order + "</span>" +
        '<div class="c-titles"><h3>' + esc(x.name) + '</h3><p class="ko">' + esc(x.nameKo) + "</p></div>" +
        '<span class="c-cat">' + esc(x.category) + "</span></div>" +
      '<p class="c-style">' + esc(x.style) + "</p>" +
      '<div class="c-tags">' + tags + "</div>" +
      '<div class="c-info">' +
        '<div class="row"><span class="k">地址</span><span class="v">' + esc(x.address) + "</span></div>" +
        '<div class="row"><span class="k">營業</span><span class="v">' + esc(x.hours) + "</span></div>" +
        '<div class="row"><span class="k">預算</span><span class="v">' + esc(x.budget) +
          ' <b class="tier">' + TIER_LABEL[x.priceTier] + "</b></span></div>" +
      "</div>" +
    "</article>"
  );
}

function buildDistrict(d) {
  const shops = sorted(d);
  return (
    '<section class="district">' +
      '<div class="d-head"><h2>' + esc(d.name) + ' <span class="ko">' + esc(d.nameKo) + "</span></h2>" +
        '<p class="tagline">' + esc(d.tagline) + "</p></div>" +
      '<p class="d-intro">' + esc(d.intro) + "</p>" +
      '<div class="d-meta"><div><b>最近車站</b>' + esc(d.nearestStation) + "</div>" +
        "<div><b>建議時段</b>" + esc(d.bestTime) + "</div></div>" +
      '<div class="map-route">' +
        '<div class="map-box">' + buildMap(d) + "</div>" +
        '<div class="route-box"><h4>步行路線建議</h4><p class="route-note">' + esc(d.routeNote) + "</p>" +
          buildRoute(d) + "</div>" +
      "</div>" +
      '<h4 class="cards-h">必逛店家</h4>' +
      '<div class="cards">' + shops.map(buildCard).join("") + "</div>" +
    "</section>"
  );
}

function buildHTML(DISTRICTS) {
  const districts = [DISTRICTS.seongsu, DISTRICTS.hannam];
  return (
    "<!DOCTYPE html><html lang='zh-Hant'><head><meta charset='utf-8'><style>" + CSS + "</style></head><body>" +
    // 封面
    '<section class="cover">' +
      '<p class="eyebrow">SEOUL · DESIGN WALK</p>' +
      "<h1>首爾質感傢俬<br>探店地圖</h1>" +
      '<p class="sub">聖水洞 성수동 × 漢南洞 한남동</p>' +
      '<p class="lead">精選必逛的選物店與傢俱陳列室，附上詳細地址、特色風格、營業時間與預估購物預算，並規劃最順暢的探店步行路線。</p>' +
      '<div class="cover-legend"><b>預算級距</b>' +
        '<span><b>₩</b> 親民（文具・雜貨・咖啡，小物 ₩10,000 起）</span>' +
        '<span><b>₩₩</b> 中階（設計選物・作家器皿，約 ₩30,000–₩300,000）</span>' +
        '<span><b>₩₩₩</b> 高階（高端設計傢俱・中古老件，可達 ₩1,000,000+）</span>' +
      "</div>" +
      '<p class="disclaimer">＊地址、營業時間與預算為整理彙編之參考值，可能變動；出發前請以 Naver Map 與店家官方 Instagram 為準。</p>' +
    "</section>" +
    districts.map(buildDistrict).join("") +
    "</body></html>"
  );
}

const CSS = `
* { box-sizing: border-box; }
body { margin: 0; font-family: "Noto Sans CJK TC", "Noto Sans CJK KR", sans-serif; color: #2b2723; line-height: 1.65; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
h1,h2,h3,h4,.serif { font-family: "Noto Serif CJK TC", "Noto Serif CJK KR", serif; }
.ko { font-family: "Noto Sans CJK KR","Noto Sans CJK TC", sans-serif; }

/* 封面 */
.cover { height: 100vh; padding: 60px 56px; display: flex; flex-direction: column; justify-content: center;
  background: linear-gradient(160deg,#f5f0e8,#efe5d4); page-break-after: always; }
.cover .eyebrow { letter-spacing: .4em; color: #b5552f; font-weight: 700; font-size: 13px; margin: 0 0 18px; }
.cover h1 { font-size: 54px; line-height: 1.15; margin: 0 0 18px; letter-spacing: .03em; }
.cover .sub { font-size: 22px; color: #6f7351; margin: 0 0 24px; font-family: "Noto Serif CJK TC", serif; }
.cover .lead { font-size: 15px; color: #6b635a; max-width: 560px; margin: 0 0 36px; }
.cover-legend { border-top: 2px solid #e3d9ca; border-bottom: 2px solid #e3d9ca; padding: 20px 0; display: grid; gap: 8px; max-width: 620px; }
.cover-legend > b { color: #a8843f; letter-spacing: .1em; }
.cover-legend span { font-size: 13.5px; color: #4a443d; }
.cover-legend span b { color: #a8843f; display: inline-block; width: 36px; }
.cover .disclaimer { margin-top: 30px; font-size: 11.5px; color: #8a8175; max-width: 600px; }

/* 區域 */
.district { padding: 40px 44px 30px; page-break-before: always; }
.d-head { display: flex; align-items: baseline; gap: 16px; border-bottom: 3px solid #b5552f; padding-bottom: 12px; }
.d-head h2 { font-size: 32px; margin: 0; }
.d-head h2 .ko { font-size: 18px; color: #6b635a; font-weight: 400; }
.d-head .tagline { color: #b5552f; font-weight: 500; margin: 0; }
.d-intro { color: #4a443d; font-size: 14px; margin: 16px 0 14px; }
.d-meta { display: grid; gap: 8px; background: #faf6ee; border: 1px solid #e3d9ca; border-radius: 12px; padding: 14px 18px; font-size: 13px; margin-bottom: 20px; }
.d-meta b { color: #6f7351; margin-right: 12px; display: inline-block; min-width: 72px; }

/* 地圖 + 路線 */
.map-route { display: grid; grid-template-columns: 1.25fr 1fr; gap: 20px; margin-bottom: 24px; page-break-inside: avoid; }
.map-box { border: 1px solid #e3d9ca; border-radius: 14px; overflow: hidden; }
svg.map { width: 100%; height: auto; display: block; }
.route-box h4 { margin: 0 0 6px; font-size: 16px; color: #2b2723; }
.route-note { font-size: 11.5px; color: #6b635a; margin: 0 0 12px; }
.route { list-style: none; margin: 0; padding: 0; }
.route li { position: relative; padding: 0 0 10px 0; display: grid; grid-template-columns: 26px 1fr; column-gap: 10px; align-items: start; }
.route .r-num { grid-row: span 2; width: 24px; height: 24px; background: #b5552f; color: #fff; border-radius: 50%; text-align: center; line-height: 24px; font-size: 12px; font-weight: 700; }
.route .r-name { font-weight: 700; font-family: "Noto Serif CJK TC", serif; font-size: 14px; }
.route .r-walk { font-size: 11.5px; color: #6f7351; }

/* 卡片 */
.cards-h { font-size: 17px; margin: 0 0 14px; border-left: 5px solid #b5552f; padding-left: 12px; }
.cards { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
.card { border: 1px solid #e3d9ca; border-radius: 14px; padding: 16px 18px; background: #fffdf9; page-break-inside: avoid; }
.c-top { display: flex; align-items: flex-start; gap: 10px; margin-bottom: 8px; }
.c-num { flex: 0 0 28px; height: 28px; background: #b5552f; color: #fff; border-radius: 50%; text-align: center; line-height: 28px; font-weight: 700; font-size: 13px; }
.c-titles h3 { font-size: 17px; margin: 0; }
.c-titles .ko { font-size: 11.5px; color: #6b635a; margin: 1px 0 0; }
.c-cat { margin-left: auto; font-size: 10.5px; padding: 3px 9px; border-radius: 999px; background: rgba(111,115,81,.14); color: #6f7351; font-weight: 700; white-space: nowrap; }
.c-style { font-size: 12px; color: #4a443d; margin: 0 0 10px; }
.c-tags { display: flex; flex-wrap: wrap; gap: 5px; margin-bottom: 10px; }
.c-tags span { font-size: 10.5px; background: rgba(181,85,47,.09); color: #b5552f; padding: 2px 8px; border-radius: 5px; }
.c-info { border-top: 1px solid #eee2d2; padding-top: 9px; display: grid; gap: 6px; }
.row { display: grid; grid-template-columns: 42px 1fr; column-gap: 8px; font-size: 11.5px; }
.row .k { color: #6f7351; font-weight: 700; }
.row .v { color: #2b2723; }
.tier { color: #a8843f; letter-spacing: 1px; }
`;

(async function () {
  if (!CHROME) {
    console.error("找不到 Chrome，請設定 CHROME_PATH 環境變數。");
    process.exit(1);
  }
  const DISTRICTS = loadData();
  const html = buildHTML(DISTRICTS);
  const browser = await puppeteer.launch({
    executablePath: CHROME,
    headless: "new",
    args: ["--no-sandbox", "--disable-setuid-sandbox"],
  });
  const page = await browser.newPage();
  await page.setContent(html, { waitUntil: "networkidle0" });
  await page.pdf({
    path: OUT,
    format: "A4",
    printBackground: true,
    margin: { top: "0", bottom: "0", left: "0", right: "0" },
  });
  await browser.close();
  console.log("PDF 已輸出：" + OUT);
})();
