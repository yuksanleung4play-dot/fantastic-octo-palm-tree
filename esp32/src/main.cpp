/**
 * HK BUS Station HUB + 餐點 + 天氣資訊板
 *
 * ESP32 firmware — polls GET /api/display/today every 60 s,
 * renders on 240×320 TFT (ILI9341 / ST7789).
 *
 * See esp32/README.md for wiring and setup.
 */

#include <Arduino.h>
#include <WiFi.h>
#include <time.h>

#include "ApiClient.h"
#include "DisplayLayout.h"
#include "DataTypes.h"
#include "config.h"

// ── Globals ──────────────────────────────────────────────────────────────────
DisplayLayout display;
DisplayData   lastGoodData;

bool     hasGoodData   = false;
bool     apiStale      = false;
uint32_t lastApiPoll   = 0;
uint32_t lastClockTick = 0;
uint32_t lastWifiRetry = 0;

// ── Wi-Fi ────────────────────────────────────────────────────────────────────
void connectWiFi() {
  if (WiFi.status() == WL_CONNECTED) return;

  Serial.printf("[WiFi] Connecting to %s …\n", WIFI_SSID);
  WiFi.mode(WIFI_STA);
  WiFi.begin(WIFI_SSID, WIFI_PASSWORD);

  const uint32_t start = millis();
  while (WiFi.status() != WL_CONNECTED && millis() - start < WIFI_RETRY_MS) {
    delay(250);
  }

  if (WiFi.status() == WL_CONNECTED) {
    Serial.printf("[WiFi] Connected — IP %s\n", WiFi.localIP().toString().c_str());
  } else {
    Serial.println("[WiFi] Connection failed — will retry");
  }
}

// ── NTP ──────────────────────────────────────────────────────────────────────
void syncNtp() {
  configTime(GMT_OFFSET_SEC, DAYLIGHT_OFFSET_SEC, NTP_SERVER);
  Serial.println("[NTP] Syncing …");

  struct tm timeinfo {};
  for (int i = 0; i < 20; ++i) {
    if (getLocalTime(&timeinfo)) {
      Serial.println("[NTP] Time synchronised");
      return;
    }
    delay(500);
  }
  Serial.println("[NTP] Sync timeout — using system clock");
}

void formatClockStrings(char *dateOut, size_t dateLen,
                        char *dayOut, size_t dayLen,
                        char *timeOut, size_t timeLen) {
  struct tm timeinfo {};
  if (!getLocalTime(&timeinfo)) {
    if (hasGoodData) {
      strncpy(dateOut, lastGoodData.date, dateLen);
      strncpy(dayOut, lastGoodData.dayOfWeek, dayLen);
      strncpy(timeOut, lastGoodData.time, timeLen);
    }
    return;
  }

  strftime(dateOut, dateLen, "%Y-%m-%d", &timeinfo);

  static const char *DAYS[] = {"SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT"};
  strncpy(dayOut, DAYS[timeinfo.tm_wday], dayLen);
  dayOut[dayLen - 1] = '\0';

  strftime(timeOut, timeLen, "%I:%M %p", &timeinfo);
  // Trim leading zero on hour: " 08:00 PM" → "08:00 PM"
  if (timeOut[0] == ' ') {
    memmove(timeOut, timeOut + 1, strlen(timeOut));
  }
}

// ── API poll ─────────────────────────────────────────────────────────────────
void pollApi() {
  DisplayData fetched {};
  const bool ok = fetchDisplayData(fetched);

  if (ok) {
    lastGoodData = fetched;
    lastGoodData.stale = apiStale;  // server-side stale flag preserved
    hasGoodData = true;
    apiStale = fetched.stale;
    display.setStaleIndicator(apiStale);
    display.drawBusSection(lastGoodData);
    display.drawWeatherSection(lastGoodData);
    display.drawMenuSection(lastGoodData);
    Serial.println("[API] Data updated");
  } else {
    apiStale = true;
    display.setStaleIndicator(true);
    Serial.println("[API] Fetch failed — keeping last data");
  }
}

// ── Clock tick (1 Hz, header only) ───────────────────────────────────────────
void tickClock() {
  char dateBuf[16], dayBuf[8], timeBuf[16];
  formatClockStrings(dateBuf, sizeof(dateBuf),
                     dayBuf, sizeof(dayBuf),
                     timeBuf, sizeof(timeBuf));
  display.drawHeaderTime(dateBuf, dayBuf, timeBuf);
}

// ── Arduino entry points ─────────────────────────────────────────────────────
void setup() {
  Serial.begin(115200);
  delay(500);
  Serial.println();
  Serial.println("=== HK Bus Display ===");

  display.begin();
  display.drawStaticChrome();

  connectWiFi();
  if (WiFi.status() == WL_CONNECTED) {
    syncNtp();
    pollApi();
    if (hasGoodData) {
      display.drawFull(lastGoodData);
    }
  }

  lastApiPoll   = millis();
  lastClockTick = millis();
  lastWifiRetry = millis();
}

void loop() {
  const uint32_t now = millis();

  // Wi-Fi reconnect
  if (WiFi.status() != WL_CONNECTED) {
    if (now - lastWifiRetry >= WIFI_RETRY_MS) {
      lastWifiRetry = now;
      connectWiFi();
      if (WiFi.status() == WL_CONNECTED) {
        syncNtp();
        pollApi();
        lastApiPoll = now;
      }
    }
  } else {
    // API poll every 60 s
    if (now - lastApiPoll >= API_POLL_MS) {
      lastApiPoll = now;
      pollApi();
    }
  }

  // Clock refresh every 1 s
  if (now - lastClockTick >= CLOCK_TICK_MS) {
    lastClockTick = now;
    tickClock();
  }
}
