#pragma once

#include "DataTypes.h"

#include <TFT_eSPI.h>

class DisplayLayout {
 public:
  void begin();
  void drawStaticChrome();
  void drawFull(const DisplayData &data);
  void drawHeaderTime(const char *dateStr, const char *dayStr, const char *timeStr);
  void drawBusSection(const DisplayData &data);
  void drawWeatherSection(const DisplayData &data);
  void drawMenuSection(const DisplayData &data);
  void drawPetAvatar();
  void setStaleIndicator(bool show);
  void setChineseEnabled(bool enabled) { chineseEnabled_ = enabled; }

  TFT_eSPI &tft() { return tft_; }

 private:
  TFT_eSPI tft_;
  bool staleVisible_ = false;
  bool chineseEnabled_ = true;
  bool petLoaded_ = false;

  uint16_t routeColor(const char *route) const;
  uint16_t weatherBg(const WeatherInfo &w) const;
  const char *weatherIconGlyph(const char *icon) const;

  void drawChinese(int x, int y, const char *text, uint16_t color, uint8_t size = 1);
  void drawAscii(int x, int y, const char *text, uint16_t color, uint8_t font = 2);
  void fillSection(int y, int h, uint16_t color);
  void drawBusRow(int row, const BusRoute &route);
  void formatDate(const char *isoDate, char *out, size_t len) const;
  void truncateDest(const char *src, char *out, size_t len) const;
};
