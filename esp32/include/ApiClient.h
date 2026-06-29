#pragma once

#include "DataTypes.h"

// Fetch /api/display/today and parse into out.
// Returns true on HTTP 200 with valid JSON; false on timeout / error.
bool fetchDisplayData(DisplayData &out);

// Build full API URL (http://host:port/path).
String buildApiUrl();
