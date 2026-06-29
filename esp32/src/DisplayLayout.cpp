#include "DisplayLayout.h"

#include "config.h"

#include <FS.h>
#include <LittleFS.h>
#include <U8g2_for_TFT_eSPI.h>

namespace {

U8g2_for_TFT_eSPI u8g2;

constexpr int STALE_DOT_X = 234;
constexpr int STALE_DOT_Y = 4;
constexpr int STALE_DOT_R = 4;

}  // namespace

void DisplayLayout::begin() {
  tft_.init();
  tft_.setRotation(0);
  tft_.fillScreen(Color::BLACK);
  tft_.setTextWrap(false);

  u8g2.begin(tft_);
  u8g2.setFontMode(1);
  u8g2.setFontDirection(0);
  u8g2.setBackgroundColor(Color::HEADER_BG);

  if (LittleFS.begin(true)) {
    petLoaded_ = LittleFS.exists(PET_BMP_PATH);
    if (!petLoaded_) {
      Serial.println("[Display] pet.bmp not found in LittleFS — skipping avatar");
    }
  } else {
    Serial.println("[Display] LittleFS mount failed");
  }
}

void DisplayLayout::fillSection(const int y, const int h, const uint16_t color) {
  tft_.fillRect(0, y, Layout::SCREEN_W, h, color);
}

void DisplayLayout::drawStaticChrome() {
  fillSection(Layout::HEADER_Y, Layout::HEADER_H, Color::HEADER_BG);
  fillSection(Layout::BUS_Y, Layout::BUS_H, Color::BUS_BG);
  fillSection(Layout::WEATHER_Y, Layout::WEATHER_H, Color::WEATHER_BG);
  fillSection(Layout::MENU_Y, Layout::MENU_H, Color::MENU_BG);
  drawPetAvatar();
}

void DisplayLayout::drawAscii(const int x, const int y, const char *text,
                              const uint16_t color, const uint8_t font) {
  tft_.setTextColor(color, Color::BLACK);
  tft_.setTextFont(font);
  tft_.setCursor(x, y);
  tft_.print(text);
}

void DisplayLayout::drawChinese(const int x, const int y, const char *text,
                                const uint16_t color, const uint8_t size) {
  if (!chineseEnabled_ || !text || text[0] == '\0') {
    drawAscii(x, y, text ? text : "", color, 2);
    return;
  }
  u8g2.setForegroundColor(color);
  u8g2.setFont(u8g2_font_wqy12_t_gb2312);
  u8g2.setCursor(x, y + 12 * size);
  u8g2.print(text);
}

void DisplayLayout::formatDate(const char *isoDate, char *out, const size_t len) const {
  // "2026-06-29" → "2026/06/29"
  if (!isoDate || strlen(isoDate) < 10) {
    snprintf(out, len, "----/--/--");
    return;
  }
  snprintf(out, len, "%.4s/%.2s/%.2s", isoDate, isoDate + 5, isoDate + 8);
}

void DisplayLayout::truncateDest(const char *src, char *out, const size_t len) const {
  if (!src) {
    out[0] = '\0';
    return;
  }
  strncpy(out, src, len - 1);
  out[len - 1] = '\0';
  // Rough CJK width limit for 240px row with route badge
  const size_t maxChars = 8;
  if (strlen(out) > maxChars) {
    strncpy(out + maxChars - 1, "…", len - maxChars);
    out[maxChars] = '\0';
  }
}

uint16_t DisplayLayout::routeColor(const char *route) const {
  if (!route) return ROUTE_COLOR_DEFAULT;
  if (strcmp(route, "38") == 0)  return ROUTE_COLOR_38;
  if (strcmp(route, "40") == 0)  return ROUTE_COLOR_40;
  if (strcmp(route, "28B") == 0) return ROUTE_COLOR_28B;
  if (strcmp(route, "13X") == 0) return ROUTE_COLOR_13X;
  return ROUTE_COLOR_DEFAULT;
}

uint16_t DisplayLayout::weatherBg(const WeatherInfo &w) const {
  if (strcmp(w.alertLevel, "black") == 0 || strcmp(w.alertLevel, "red") == 0 ||
      strcmp(w.alertLevel, "typhoon") == 0) {
    return Color::ALERT_RED;
  }
  if (strcmp(w.alertLevel, "yellow") == 0) {
    return Color::ALERT_YELLOW;
  }
  return Color::WEATHER_BG;
}

const char *DisplayLayout::weatherIconGlyph(const char *icon) const {
  if (!icon) return "*";
  if (strcmp(icon, "sunny") == 0 || strcmp(icon, "sunny_morning") == 0) return "#";
  if (strcmp(icon, "cloudy_sunny") == 0) return "~";
  if (strcmp(icon, "cloudy") == 0) return "=";
  if (strcmp(icon, "rainy") == 0 || strcmp(icon, "heavy_rain") == 0) return "/";
  if (strcmp(icon, "thunderstorm") == 0) return "!";
  if (strcmp(icon, "hot") == 0) return "H";
  if (strcmp(icon, "warm") == 0) return "W";
  if (strcmp(icon, "cool") == 0) return "C";
  if (strcmp(icon, "cold") == 0) return "c";
  return "*";
}

void DisplayLayout::setStaleIndicator(const bool show) {
  if (show == staleVisible_) return;
  staleVisible_ = show;
  if (show) {
    tft_.fillCircle(STALE_DOT_X, STALE_DOT_Y, STALE_DOT_R, Color::STALE_DOT);
  } else {
    tft_.fillCircle(STALE_DOT_X, STALE_DOT_Y, STALE_DOT_R, Color::HEADER_BG);
  }
}

void DisplayLayout::drawPetAvatar() {
  if (!petLoaded_) return;
  File f = LittleFS.open(PET_BMP_PATH, "r");
  if (!f) return;

  // BMP header: 54 bytes, then RGB565 pixel data (little-endian)
  f.seek(54);
  const size_t pixelBytes = PET_AVATAR_W * PET_AVATAR_H * 2;
  uint16_t *buf = static_cast<uint16_t *>(malloc(pixelBytes));
  if (!buf) {
    f.close();
    return;
  }
  f.read(reinterpret_cast<uint8_t *>(buf), pixelBytes);
  f.close();

  tft_.pushImage(PET_AVATAR_X, PET_AVATAR_Y, PET_AVATAR_W, PET_AVATAR_H, buf);
  free(buf);
}

void DisplayLayout::drawHeaderTime(const char *dateStr, const char *dayStr,
                                   const char *timeStr) {
  // Only redraw the left portion of the header (preserve avatar area)
  tft_.fillRect(0, 0, PET_AVATAR_X - 2, Layout::HEADER_H, Color::HEADER_BG);

  char formatted[16];
  formatDate(dateStr, formatted, sizeof(formatted));

  drawAscii(6, 8, formatted, Color::WHITE, 2);
  drawAscii(130, 8, dayStr ? dayStr : "", Color::WHITE, 2);
  drawAscii(6, 32, timeStr ? timeStr : "", Color::WHITE, 4);

  if (staleVisible_) {
    tft_.fillCircle(STALE_DOT_X, STALE_DOT_Y, STALE_DOT_R, Color::STALE_DOT);
  }
}

void DisplayLayout::drawBusRow(const int row, const BusRoute &route) {
  const int y = Layout::BUS_Y + row * Layout::BUS_ROW_H;
  tft_.fillRect(0, y, Layout::SCREEN_W, Layout::BUS_ROW_H, Color::BUS_BG);

  // Route badge
  const uint16_t badgeColor = routeColor(route.route);
  tft_.fillRoundRect(4, y + 8, 36, 22, 4, badgeColor);
  drawAscii(10, y + 12, route.route, Color::WHITE, 2);

  // Destination (Chinese)
  char destShort[32];
  truncateDest(route.dest, destShort, sizeof(destShort));
  u8g2.setBackgroundColor(Color::BUS_BG);
  drawChinese(46, y + 10, destShort, Color::WHITE);

  // ETA or remark
  const int rightX = 148;
  if (route.remark[0] != '\0') {
    drawAscii(rightX, y + 14, route.remark, Color::WHITE, 2);
  } else if (route.etaCount > 0) {
    char etaBuf[20];
    if (route.etaCount == 1) {
      snprintf(etaBuf, sizeof(etaBuf), "%d min", route.eta[0]);
    } else {
      snprintf(etaBuf, sizeof(etaBuf), "%d/%d", route.eta[0], route.eta[1]);
    }
    drawAscii(rightX, y + 14, etaBuf, Color::WHITE, 2);
  } else {
    drawAscii(rightX, y + 14, "--", Color::WHITE, 2);
  }
}

void DisplayLayout::drawBusSection(const DisplayData &data) {
  for (int i = 0; i < Layout::MAX_BUS_ROUTES; ++i) {
    if (i < data.busCount) {
      drawBusRow(i, data.bus[i]);
    } else {
      tft_.fillRect(0, Layout::BUS_Y + i * Layout::BUS_ROW_H,
                    Layout::SCREEN_W, Layout::BUS_ROW_H, Color::BUS_BG);
    }
  }
}

void DisplayLayout::drawWeatherSection(const DisplayData &data) {
  const WeatherInfo &w = data.weather;
  const uint16_t bg = weatherBg(w);
  const uint16_t fg =
      (bg == Color::ALERT_YELLOW) ? Color::BLACK : Color::WHITE;

  fillSection(Layout::WEATHER_Y, Layout::WEATHER_H, bg);

  char line1[48];
  snprintf(line1, sizeof(line1), "%dC  %s  Hum:%d%%",
           w.temperature, weatherIconGlyph(w.icon), w.humidity);
  drawAscii(6, Layout::WEATHER_Y + 8, line1, fg, 2);

  char line2[48];
  snprintf(line2, sizeof(line2), "UV:%d  %s", w.uvIndex, w.alertText);
  drawAscii(6, Layout::WEATHER_Y + 32, line2, fg, 2);
}

void DisplayLayout::drawMenuSection(const DisplayData &data) {
  fillSection(Layout::MENU_Y, Layout::MENU_H, Color::MENU_BG);
  u8g2.setBackgroundColor(Color::MENU_BG);

  if (data.menuPeriodCount == 0) {
    drawChinese(6, Layout::MENU_Y + 6, "今日無餐單", Color::WHITE);
    return;
  }

  int lineY = Layout::MENU_Y + 4;
  for (int p = 0; p < data.menuPeriodCount; ++p) {
    const MenuPeriod &mp = data.menu[p];
    if (mp.itemCount == 0) continue;

    char line[64];
    snprintf(line, sizeof(line), "%s:", mp.periodLabel);
    for (int i = 0; i < mp.itemCount && i < 2; ++i) {
      const size_t used = strlen(line);
      const char *suffix = (mp.items[i].stock[0] && strcmp(mp.items[i].stock, "normal") != 0)
                               ? "(!)"
                               : "";
      snprintf(line + used, sizeof(line) - used, "%s%s%s",
               (used > 0 && line[used - 1] != ':') ? " " : "",
               mp.items[i].name, suffix);
    }
    drawChinese(6, lineY, line, Color::WHITE);
    lineY += 18;
    if (lineY > Layout::MENU_Y + Layout::MENU_H - 14) break;
  }
}

void DisplayLayout::drawFull(const DisplayData &data) {
  drawStaticChrome();

  char timeBuf[16];
  snprintf(timeBuf, sizeof(timeBuf), "%s", data.time);
  drawHeaderTime(data.date, data.dayOfWeek, timeBuf);

  drawBusSection(data);
  drawWeatherSection(data);
  drawMenuSection(data);
  setStaleIndicator(data.stale);
}
