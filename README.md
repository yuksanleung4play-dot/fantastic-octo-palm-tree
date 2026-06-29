# HK BUS Station HUB — 資訊板系統

香港巴士站資訊板 + 餐點 + 天氣顯示專案。

## 組件

| 目錄 | 說明 |
|------|------|
| [`esp32/`](esp32/) | ESP32 韌體（240×320 TFT 顯示端） |

## ESP32 韌體

請參閱 [`esp32/README.md`](esp32/README.md) 了解硬體接線、設定與燒錄步驟。

```bash
cd esp32
# 編輯 include/config.h 設定 Wi-Fi 與 API
pio run -t upload
```

## 後端 API

後端 Node.js `GET /api/display/today` 端點規格見需求文件（`bus-display-requirements.md`）。
