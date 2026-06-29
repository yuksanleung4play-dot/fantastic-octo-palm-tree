#pragma once

// Development defaults — replace with your real values before flashing.
// For production, copy config.example.h → config.h and edit.

#define WIFI_SSID     "YOUR_WIFI_SSID"
#define WIFI_PASSWORD "YOUR_WIFI_PASSWORD"

#define API_HOST      "192.168.1.100"
#define API_PORT      3000
#define API_PATH      "/api/display/today"
#define DISPLAY_TOKEN "your-display-bearer-token"

#define NTP_SERVER    "pool.ntp.org"
#define GMT_OFFSET_SEC      (8 * 3600)
#define DAYLIGHT_OFFSET_SEC 0

#define WIFI_RETRY_MS    5000
#define API_POLL_MS      60000
#define HTTP_TIMEOUT_MS  5000
#define CLOCK_TICK_MS    1000

#define PET_BMP_PATH     "/pet.bmp"
#define PET_AVATAR_X     192
#define PET_AVATAR_Y     6
#define PET_AVATAR_W     48
#define PET_AVATAR_H     48

#define ROUTE_COLOR_38   0x07E0
#define ROUTE_COLOR_40   0xFD20
#define ROUTE_COLOR_28B  0xF81F
#define ROUTE_COLOR_13X  0x07FF
#define ROUTE_COLOR_DEFAULT 0xFFFF
