#pragma once

#include <Arduino.h>

// ── Colour constants (RGB565) ───────────────────────────────────────────────
namespace Color {
constexpr uint16_t WHITE      = 0xFFFF;
constexpr uint16_t BLACK      = 0x0000;
constexpr uint16_t HEADER_BG  = 0xF800;  // red
constexpr uint16_t BUS_BG     = 0x0010;  // dark blue
constexpr uint16_t WEATHER_BG = 0x001F;  // blue
constexpr uint16_t MENU_BG    = 0x2104;  // dark grey
constexpr uint16_t ALERT_RED  = 0xF800;
constexpr uint16_t ALERT_YELLOW = 0xFFE0;
constexpr uint16_t STALE_DOT  = 0xF800;
}

// ── Layout (240 × 320) ────────────────────────────────────────────────────────
namespace Layout {
constexpr int SCREEN_W  = 240;
constexpr int SCREEN_H  = 320;
constexpr int HEADER_Y  = 0;
constexpr int HEADER_H  = 60;
constexpr int BUS_Y     = 60;
constexpr int BUS_H     = 160;
constexpr int WEATHER_Y = 220;
constexpr int WEATHER_H = 60;
constexpr int MENU_Y    = 280;
constexpr int MENU_H    = 40;
constexpr int BUS_ROW_H   = 40;
constexpr int MAX_BUS_ROUTES = 4;
constexpr int MAX_MENU_PERIODS = 2;
constexpr int MAX_MENU_ITEMS   = 8;
constexpr int MAX_ETA          = 2;
}

struct MenuItem {
  char name[32];
  char stock[12];  // "normal" | "low" | "sold_out"
};

struct MenuPeriod {
  char period[12];       // "lunch" | "dinner"
  char periodLabel[16];  // "午餐" | "晚餐"
  MenuItem items[Layout::MAX_MENU_ITEMS];
  int itemCount = 0;
};

struct BusRoute {
  char route[8];
  char dest[32];
  int  eta[Layout::MAX_ETA];
  int  etaCount = 0;
  char remark[24];  // "Scheduled Bus", "HOLD", etc. — empty = null
};

struct WeatherInfo {
  int  temperature = 0;
  int  humidity    = 0;
  int  uvIndex     = 0;
  char icon[20];         // "cloudy", "sunny", …
  char alertLevel[12];   // "normal" | "yellow" | "red" | "black" | "typhoon"
  char alertText[48];
};

struct PetInfo {
  char name[24];
  char mood[16];
  bool hasAvatar = false;
};

struct DisplayData {
  char date[12];        // "2026-06-29"
  char dayOfWeek[4];    // "MON"
  char time[8];         // "20:00" from API (display uses NTP)
  MenuPeriod menu[Layout::MAX_MENU_PERIODS];
  int menuPeriodCount = 0;
  BusRoute bus[Layout::MAX_BUS_ROUTES];
  int busCount = 0;
  WeatherInfo weather;
  PetInfo pet;
  bool stale = false;
  bool valid = false;
};
