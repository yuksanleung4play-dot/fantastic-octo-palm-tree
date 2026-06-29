# ESP32 資訊板韌體

HK BUS Station HUB + 餐點 + 天氣資訊板 — 240×320 TFT 顯示端韌體。

## 功能

- Wi-Fi 連線（失敗每 5 秒重試）
- NTP 時間同步（`pool.ntp.org`，Asia/Hong_Kong UTC+8）
- 每 60 秒輪詢 `GET /api/display/today`（Bearer Token）
- HTTP 逾時 5 秒；失敗時保留上次資料，右上角紅點警示
- 分區渲染：標題欄 / 巴士 ETA / 天氣 / 餐點
- 寵物頭像：SPIFFS/LittleFS 靜態 BMP（48×48）
- 中文顯示：U8g2_for_TFT_eSPI（文泉驛 12pt）

## 硬體需求

| 項目 | 規格 |
|------|------|
| MCU | ESP32（建議 ESP32-WROOM） |
| 螢幕 | 2.4–3.2" SPI TFT，ILI9341 或 ST7789，240×320 |
| 連線 | Wi-Fi（2.4 GHz） |

### 預設 SPI 腳位（可在 `platformio.ini` 修改）

| 信號 | GPIO |
|------|------|
| MOSI | 23 |
| MISO | 19 |
| SCLK | 18 |
| CS   | 5  |
| DC   | 2  |
| RST  | 4  |

若使用 ST7789，將 `platformio.ini` 中 `-DILI9341_DRIVER=1` 改為 `-DST7789_DRIVER=1`。

## 快速開始

### 1. 設定連線參數

編輯 `include/config.h`（或從 `include/config.example.h` 複製）：

```cpp
#define WIFI_SSID     "你的WiFi"
#define WIFI_PASSWORD "你的密碼"
#define API_HOST      "192.168.1.100"   // NAS IP
#define API_PORT      3000
#define DISPLAY_TOKEN "your-token"
```

### 2. 準備寵物頭像（選用）

1. 裁切正方形照片為 **48×48 px**
2. 轉為 **RGB565 BMP**（可用 [image2cpp](https://lvgl.io/tools/imageconverter) 或 GIMP）
3. 存為 `data/pet.bmp`
4. 上傳檔案系統：

```bash
pio run -t uploadfs
```

### 3. 編譯與燒錄

```bash
cd esp32
pio run -t upload
pio device monitor
```

## 畫面佈局（240×320）

```
┌───────────────────┬────────┐  標題欄 60px（紅底）
│  2026/06/29   MON │ avatar │  寵物頭像 48×48 @ (192,6)
│  08:00 PM         │  48px  │
├───────────────────┴────────┤  巴士區 160px（深藍底）
│ 38  觀塘（東）       5/14  │
│ 40  悅來花園        14/38  │
│ …                          │
├────────────────────────────┤  天氣區 60px（藍底）
│ 30C  =  Hum:82%            │
│ UV:1  Normal Weather       │
├────────────────────────────┤  餐點區 40px（深灰底）
│ 晚餐: 豬扒飯 沙嗲牛肉…      │
└────────────────────────────┘
```

## API 回應格式

韌體預期後端回傳如下 JSON（詳見專案需求文件）：

```json
{
  "date": "2026-06-29",
  "day_of_week": "MON",
  "time": "20:00",
  "menu": [{ "period": "dinner", "period_label": "晚餐", "items": [...] }],
  "bus": [{ "route": "38", "dest": "觀塘（東）", "eta": [5, 18], "remark": null }],
  "weather": { "temperature": 30, "humidity": 82, "uv_index": 1, "icon": "cloudy", "alert_level": "normal", "alert_text": "Normal Weather" },
  "pet": { "name": "...", "avatar_url": "/assets/pet.bmp", "mood": "happy" }
}
```

## 專案結構

```
esp32/
├── platformio.ini          # 依賴與 TFT 腳位
├── include/
│   ├── config.h            # 連線設定（請自行修改）
│   ├── config.example.h
│   ├── DataTypes.h         # 資料結構
│   ├── ApiClient.h         # HTTP + JSON 解析
│   └── DisplayLayout.h     # 畫面渲染
├── src/
│   ├── main.cpp            # 主迴圈
│   ├── ApiClient.cpp
│   └── DisplayLayout.cpp
├── data/
│   └── pet.bmp             # 寵物頭像（自行加入）
└── tools/
    └── gen_placeholder_bmp.py
```

## 疑難排解

| 問題 | 解法 |
|------|------|
| 螢幕全白/全黑 | 檢查 `platformio.ini` 驅動型號與 SPI 腳位 |
| 中文亂碼/方塊 | 確認已安裝 `U8g2_for_TFT_eSPI`；UTF-8 字串需為 GB2312 相容繁簡中文 |
| API 連不上 | 確認 NAS IP、Token、ESP32 與 NAS 在同一 LAN |
| 紅點一直亮 | 上次 API 請求失敗；檢查 Serial Monitor 日誌 |
| 頭像不顯示 | 執行 `pio run -t uploadfs`，確認 `data/pet.bmp` 為 48×48 RGB565 BMP |

## 授權

與上層專案相同。
