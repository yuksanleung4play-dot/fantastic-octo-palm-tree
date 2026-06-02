# 首爾質感傢俬探店地圖（PDF）｜ 聖水洞 × 漢南洞

一份可直接在手機 / GitHub App 預覽的 **PDF 探店指南**,整理首爾 **聖水洞 (성수동)** 與 **漢南洞 (한남동)** 的質感傢俬店、選物店與傢俱陳列室。

## 直接看 PDF

打開 repo 裡的 **`首爾質感傢俬探店地圖.pdf`** 即可(GitHub iOS App / 網頁都能直接預覽,也可下載列印或存進手機)。

## PDF 內容（A4・5 頁）

- **封面**:主題、預算級距總覽與免責聲明。
- **每個區域各一段**(聖水洞 7 間、漢南洞 6 間):
  - 區域簡介、最近車站、建議時段
  - **示意探店地圖**:標記順序編號與步行路線
  - **步行路線建議**:依最順暢動線排序,標示各站之間概略步行時間
  - **必逛店家卡片**,每間含:
    - 詳細地址(中文 + 韓文原文,方便貼進 Naver Map)
    - 特色風格描述與亮點標籤
    - 營業時間
    - 預估購物預算與級距（₩ / ₩₩ / ₩₩₩）

## 收錄店家

- **聖水洞**:LCDC SEOUL、Point of View、Object 聖水、MTL 聖水、聖水聯邦、聖水中古傢俱街、大林倉庫藝廊
- **漢南洞**:TWL、Chapter1 漢南、MMMG 漢南、D&DEPARTMENT SEOUL、H PIX 漢南、Object 漢南

## 想修改內容、重新產生 PDF？

資料集中在 `data.js`,改完後重新跑產生器即可:

```bash
# 需求:Node.js、系統已安裝 Google Chrome 與 Noto CJK 字型
#   字型安裝(Debian/Ubuntu):sudo apt-get install -y fonts-noto-cjk
npm install puppeteer-core
node build-pdf.js
# 產出:首爾質感傢俬探店地圖.pdf
```

`data.js` 每間店欄位:`order`(順序/地圖編號)、`name` / `nameKo`(中/韓店名)、
`category`(選物店 / 傢俱陳列室 / 複合空間)、`address` / `style` / `hours` / `budget`、
`highlights`(亮點標籤)、`priceTier`(1 親民 / 2 中階 / 3 高階)、`x` / `y`(示意地圖座標)。

## 檔案結構

| 檔案 | 說明 |
| --- | --- |
| `首爾質感傢俬探店地圖.pdf` | 最終成品,直接開來看 |
| `data.js` | 兩區店家與路線資料 |
| `build-pdf.js` | 以 Chrome 將資料排版輸出 PDF 的產生器 |

## 免責聲明

地址、營業時間與預算為整理彙編之**參考值**,可能變動。出發前請以 **Naver Map** 與店家官方 **Instagram** 為準。
