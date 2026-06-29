#include "ApiClient.h"

#include "config.h"

#include <ArduinoJson.h>
#include <HTTPClient.h>
#include <WiFi.h>

namespace {

void copyStr(char *dst, size_t len, const char *src) {
  if (!src) {
    dst[0] = '\0';
    return;
  }
  strncpy(dst, src, len - 1);
  dst[len - 1] = '\0';
}

void parseMenu(const JsonArray &arr, DisplayData &out) {
  out.menuPeriodCount = 0;
  for (JsonObject period : arr) {
    if (out.menuPeriodCount >= Layout::MAX_MENU_PERIODS) break;
    MenuPeriod &mp = out.menu[out.menuPeriodCount++];
    copyStr(mp.period, sizeof(mp.period), period["period"] | "");
    copyStr(mp.periodLabel, sizeof(mp.periodLabel), period["period_label"] | "");
    mp.itemCount = 0;
    JsonArray items = period["items"].as<JsonArray>();
    for (JsonObject item : items) {
      if (mp.itemCount >= Layout::MAX_MENU_ITEMS) break;
      const char *stock = item["stock"] | "normal";
      if (strcmp(stock, "sold_out") == 0) continue;
      MenuItem &mi = mp.items[mp.itemCount++];
      copyStr(mi.name, sizeof(mi.name), item["name"] | "");
      copyStr(mi.stock, sizeof(mi.stock), stock);
    }
  }
}

void parseBus(const JsonArray &arr, DisplayData &out) {
  out.busCount = 0;
  for (JsonObject route : arr) {
    if (out.busCount >= Layout::MAX_BUS_ROUTES) break;
    BusRoute &br = out.bus[out.busCount++];
    copyStr(br.route, sizeof(br.route), route["route"] | "");
    copyStr(br.dest, sizeof(br.dest), route["dest"] | "");
    br.etaCount = 0;
    if (route["eta"].is<JsonArray>()) {
      for (int eta : route["eta"].as<JsonArray>()) {
        if (br.etaCount >= Layout::MAX_ETA) break;
        br.eta[br.etaCount++] = eta;
      }
    }
    if (route["remark"].isNull()) {
      br.remark[0] = '\0';
    } else {
      copyStr(br.remark, sizeof(br.remark), route["remark"] | "");
    }
  }
}

void parseWeather(const JsonObject &w, DisplayData &out) {
  WeatherInfo &wi = out.weather;
  wi.temperature = w["temperature"] | 0;
  wi.humidity    = w["humidity"] | 0;
  wi.uvIndex     = w["uv_index"] | 0;
  copyStr(wi.icon, sizeof(wi.icon), w["icon"] | "cloudy");
  copyStr(wi.alertLevel, sizeof(wi.alertLevel), w["alert_level"] | "normal");
  copyStr(wi.alertText, sizeof(wi.alertText), w["alert_text"] | "Normal");
}

void parsePet(const JsonObject &p, DisplayData &out) {
  copyStr(out.pet.name, sizeof(out.pet.name), p["name"] | "");
  copyStr(out.pet.mood, sizeof(out.pet.mood), p["mood"] | "");
  out.pet.hasAvatar = (p["avatar_url"] | nullptr) != nullptr;
}

bool parsePayload(const String &payload, DisplayData &out) {
  JsonDocument doc;
  DeserializationError err = deserializeJson(doc, payload);
  if (err) {
    Serial.printf("[API] JSON parse error: %s\n", err.c_str());
    return false;
  }

  copyStr(out.date, sizeof(out.date), doc["date"] | "");
  copyStr(out.dayOfWeek, sizeof(out.dayOfWeek), doc["day_of_week"] | "");
  copyStr(out.time, sizeof(out.time), doc["time"] | "");
  out.stale = doc["stale"] | false;

  if (doc["menu"].is<JsonArray>()) {
    parseMenu(doc["menu"].as<JsonArray>(), out);
  } else {
    out.menuPeriodCount = 0;
  }

  if (doc["bus"].is<JsonArray>()) {
    parseBus(doc["bus"].as<JsonArray>(), out);
  } else {
    out.busCount = 0;
  }

  if (doc["weather"].is<JsonObject>()) {
    parseWeather(doc["weather"].as<JsonObject>(), out);
  }

  if (doc["pet"].is<JsonObject>()) {
    parsePet(doc["pet"].as<JsonObject>(), out);
  }

  out.valid = true;
  return true;
}

}  // namespace

String buildApiUrl() {
  return String("http://") + API_HOST + ":" + String(API_PORT) + API_PATH;
}

bool fetchDisplayData(DisplayData &out) {
  if (WiFi.status() != WL_CONNECTED) {
    Serial.println("[API] Wi-Fi not connected");
    return false;
  }

  HTTPClient http;
  const String url = buildApiUrl();
  http.setTimeout(HTTP_TIMEOUT_MS);
  http.begin(url);
  http.addHeader("Authorization", String("Bearer ") + DISPLAY_TOKEN);
  http.addHeader("Accept", "application/json");

  Serial.printf("[API] GET %s\n", url.c_str());
  const int code = http.GET();
  String payload = http.getString();
  http.end();

  if (code != HTTP_CODE_OK) {
    Serial.printf("[API] HTTP %d\n", code);
    return false;
  }

  return parsePayload(payload, out);
}
