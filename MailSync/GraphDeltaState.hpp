//
//  GraphDeltaState.hpp
//  MailSync
//
//  Pure helpers for the Microsoft Graph calendar sync: the per-calendar sync token, the
//  full-resync decision, URL/time formatting, and mapping of Graph calendar colors. They hold
//  no state and touch neither the network nor the database so they can be unit tested.
//
//  Use of this file is subject to the terms and conditions defined
//  in 'LICENSE.md', which is part of the Mailspring-Sync package.
//

#ifndef GraphDeltaState_hpp
#define GraphDeltaState_hpp

#include <cctype>
#include <cstdio>
#include <ctime>
#include <string>
#include <stdexcept>
#include "json.hpp"

// A calendarView/delta link is bound to the calendar and the date window it was created
// for, so the window is stored next to the link in Calendar.syncToken.
struct GraphDeltaState {
    std::string deltaLink;
    time_t windowStart = 0;
    time_t windowEnd = 0;
};

// The sync window start moves forward with the clock; a link whose window starts more than
// this long before the desired start is discarded so events that aged out get dropped.
static const time_t GRAPH_MAX_WINDOW_LAG_SEC = 7 * 24 * 60 * 60;

inline std::string GraphDeltaStateToToken(const GraphDeltaState & state) {
    nlohmann::json j = {
        {"deltaLink", state.deltaLink},
        {"windowStart", (long long)state.windowStart},
        {"windowEnd", (long long)state.windowEnd},
    };
    return j.dump();
}

// Returns false for an empty token, a token written by another provider's sync (CalDAV
// sync-tokens are URIs, not JSON objects), or one missing a field.
inline bool GraphDeltaStateFromToken(const std::string & token, GraphDeltaState & out) {
    if (token.empty() || token[0] != '{') {
        return false;
    }
    try {
        nlohmann::json j = nlohmann::json::parse(token);
        if (!j.is_object() || !j["deltaLink"].is_string() || !j["windowStart"].is_number_integer() || !j["windowEnd"].is_number_integer()) {
            return false;
        }
        GraphDeltaState parsed;
        parsed.deltaLink = j["deltaLink"].get<std::string>();
        parsed.windowStart = (time_t)j["windowStart"].get<long long>();
        parsed.windowEnd = (time_t)j["windowEnd"].get<long long>();
        if (parsed.deltaLink.empty()) {
            return false;
        }
        out = parsed;
        return true;
    } catch (nlohmann::json::exception &) {
        return false;
    }
}

// True when the stored token cannot be continued incrementally: none/unusable, or its
// window start lags the desired start by more than GRAPH_MAX_WINDOW_LAG_SEC. (A 410 from the
// server is handled by the caller, which only learns of it from the response.)
inline bool GraphNeedsFullResync(const std::string & token, time_t desiredWindowStart) {
    GraphDeltaState state;
    if (!GraphDeltaStateFromToken(token, state)) {
        return true;
    }
    return desiredWindowStart - state.windowStart > GRAPH_MAX_WINDOW_LAG_SEC;
}

// ISO 8601 UTC, e.g. 2025-01-15T09:30:00Z, as Graph expects in startDateTime/endDateTime.
inline std::string GraphFormatISO8601UTC(time_t t) {
    char buffer[24];
    tm parts {};
#if defined(_MSC_VER)
    gmtime_s(&parts, &t);
#else
    gmtime_r(&t, &parts);
#endif
    strftime(buffer, sizeof(buffer), "%Y-%m-%dT%H:%M:%SZ", &parts);
    return std::string(buffer);
}

// Percent-encodes everything outside the RFC 3986 unreserved set. Graph ids are base64-like
// and can contain '=', '+' or '/'.
inline std::string GraphUrlEncode(const std::string & value) {
    static const char * hex = "0123456789ABCDEF";
    std::string out;
    for (unsigned char c : value) {
        if (isalnum(c) || c == '-' || c == '_' || c == '.' || c == '~') {
            out += (char)c;
        } else {
            out += '%';
            out += hex[c >> 4];
            out += hex[c & 15];
        }
    }
    return out;
}

inline std::string GraphDeltaStartURL(const std::string & graphCalendarId, time_t windowStart, time_t windowEnd) {
    return "https://graph.microsoft.com/v1.0/me/calendars/" + GraphUrlEncode(graphCalendarId) +
           "/calendarView/delta?startDateTime=" + GraphUrlEncode(GraphFormatISO8601UTC(windowStart)) +
           "&endDateTime=" + GraphUrlEncode(GraphFormatISO8601UTC(windowEnd));
}

// Seconds to wait before retrying a throttled (429/503) request: the server's delay-seconds
// Retry-After, clamped so a hostile or buggy value can never park the sync thread for long.
inline int GraphBoundedRetryAfterSec(const std::string & retryAfter, int fallbackSec = 5, int maxSec = 60) {
    int seconds = fallbackSec;
    if (!retryAfter.empty() && isdigit((unsigned char)retryAfter[0])) {
        try {
            seconds = std::stoi(retryAfter);
        } catch (std::out_of_range &) {
            seconds = maxSec;
        } catch (...) {
            seconds = fallbackSec;
        }
    }
    if (seconds < 1) seconds = 1;
    if (seconds > maxSec) seconds = maxSec;
    return seconds;
}

// Graph's `color` enum (used when a calendar has no hexColor) mapped to the hex values
// Outlook renders them with.
inline std::string GraphCalendarColor(const std::string & hexColor, const std::string & colorName) {
    if (hexColor.size() >= 4 && hexColor[0] == '#') {
        return hexColor;
    }
    static const struct { const char * name; const char * hex; } table[] = {
        {"lightBlue", "#6ba4e7"}, {"lightGreen", "#7fba00"}, {"lightOrange", "#f7a35c"},
        {"lightGray", "#a8a8a8"}, {"lightYellow", "#f5d300"}, {"lightTeal", "#4ec2c2"},
        {"lightPink", "#f08bb4"}, {"lightBrown", "#b08968"}, {"lightRed", "#e8736c"},
        {"maxColor", "#4285f4"},
    };
    for (const auto & entry : table) {
        if (colorName == entry.name) return entry.hex;
    }
    return "#4285f4";
}

#endif /* GraphDeltaState_hpp */
