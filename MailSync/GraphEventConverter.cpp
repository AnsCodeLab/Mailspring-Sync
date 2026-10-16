//
//  GraphEventConverter.cpp
//  MailSync
//
//  Use of this file is subject to the terms and conditions defined
//  in 'LICENSE.md', which is part of the Mailspring-Sync package.
//

#include "GraphEventConverter.hpp"

#include <algorithm>
#include <cctype>
#include <ctime>
#include <stdexcept>
#include <vector>

#ifdef _MSC_VER
#define timegm _mkgmtime
#endif

using namespace std;
using namespace nlohmann;

static const size_t ICS_MAX_LINE_OCTETS = 75;

string ICSEscapeText(const string & text) {
    string out;
    out.reserve(text.size());
    for (size_t i = 0; i < text.size(); i++) {
        unsigned char c = text[i];
        switch (c) {
            case '\\': out += "\\\\"; break;
            case ';': out += "\\;"; break;
            case ',': out += "\\,"; break;
            case '\r':
                if (i + 1 < text.size() && text[i + 1] == '\n') i++;
                out += "\\n";
                break;
            case '\n': out += "\\n"; break;
            default:
                // Other control characters are not representable in a TEXT value.
                if (c >= 0x20 || c == '\t') out += (char)c;
        }
    }
    return out;
}

string ICSFoldLine(const string & line) {
    string out;
    size_t pos = 0;
    bool first = true;
    do {
        // Continuation lines spend one octet on the leading space.
        size_t room = first ? ICS_MAX_LINE_OCTETS : ICS_MAX_LINE_OCTETS - 1;
        size_t end = min(line.size(), pos + room);
        if (end < line.size()) {
            // Don't split inside a multi-byte UTF-8 sequence (continuation bytes are 10xxxxxx).
            while (end > pos + 1 && (static_cast<unsigned char>(line[end]) & 0xC0) == 0x80) {
                end--;
            }
        }
        if (!first) out += ' ';
        out.append(line, pos, end - pos);
        out += "\r\n";
        pos = end;
        first = false;
    } while (pos < line.size());
    return out;
}

string GraphEventUID(const string & graphEventId) {
    return graphEventId + "@graph.microsoft.com";
}

string ICSStripCR(const string & ics) {
    string out;
    out.reserve(ics.size());
    for (char c : ics) {
        if (c != '\r') out += c;
    }
    return out;
}

static string stringField(const json & j, const char * key) {
    if (!j.is_object()) return "";
    auto it = j.find(key);
    if (it == j.end() || !it->is_string()) return "";
    return it->get<string>();
}

static bool boolField(const json & j, const char * key) {
    if (!j.is_object()) return false;
    auto it = j.find(key);
    return it != j.end() && it->is_boolean() && it->get<bool>();
}

static const json * objectField(const json & j, const char * key) {
    if (!j.is_object()) return nullptr;
    auto it = j.find(key);
    if (it == j.end() || !it->is_object()) return nullptr;
    return &*it;
}

static string lowercase(string s) {
    transform(s.begin(), s.end(), s.begin(), [](unsigned char c) { return (char)tolower(c); });
    return s;
}

// Parses "2025-01-15T10:30:00.0000000" (any fractional part, optional trailing Z) into
// "20250115" and "103000".
static bool parseGraphDateTime(const string & text, string & date, string & time) {
    if (text.size() < 19 || text[4] != '-' || text[7] != '-' || (text[10] != 'T' && text[10] != ' ') || text[13] != ':' || text[16] != ':') {
        return false;
    }
    const size_t digitAt[] = {0, 1, 2, 3, 5, 6, 8, 9, 11, 12, 14, 15, 17, 18};
    for (size_t i : digitAt) {
        if (!isdigit((unsigned char)text[i])) return false;
    }
    date = text.substr(0, 4) + text.substr(5, 2) + text.substr(8, 2);
    time = text.substr(11, 2) + text.substr(14, 2) + text.substr(17, 2);
    return true;
}

// Graph reports UTC as "UTC" (and, depending on API surface, one of the other spellings below).
// Everything else, including "GMT Standard Time" (London, which observes DST), needs time zone
// rules this converter does not have.
static bool isUTCZone(const string & zone) {
    if (zone.empty()) return true;
    string z = lowercase(zone);
    return z == "utc" || z == "etc/utc" || z == "tzone://microsoft/utc" || z == "greenwich standard time";
}

// Seconds since the epoch for a Graph UTC date-time, independent of the process time zone.
static bool utcUnix(const string & text, long long & out) {
    string date, time;
    if (!parseGraphDateTime(text, date, time)) return false;
    tm parts {};
    parts.tm_year = stoi(date.substr(0, 4)) - 1900;
    parts.tm_mon = stoi(date.substr(4, 2)) - 1;
    parts.tm_mday = stoi(date.substr(6, 2));
    parts.tm_hour = stoi(time.substr(0, 2));
    parts.tm_min = stoi(time.substr(2, 2));
    parts.tm_sec = stoi(time.substr(4, 2));
    out = (long long)timegm(&parts);
    return true;
}

// Returns the ICS value for a Graph dateTimeTimeZone, e.g. 20250115T103000Z or, for all-day
// events, 20250115.
static string icsDateTime(const json & dtz, bool dateOnly, const char * what) {
    string date, time;
    if (!parseGraphDateTime(stringField(dtz, "dateTime"), date, time)) {
        throw invalid_argument(string("Graph event has no usable ") + what);
    }
    if (dateOnly) return date;
    if (!isUTCZone(stringField(dtz, "timeZone"))) {
        throw invalid_argument(string("Graph event ") + what + " is not in UTC (" + stringField(dtz, "timeZone") + ")");
    }
    return date + "T" + time + "Z";
}

// Parses Graph's ISO 8601 `Z` timestamps (lastModifiedDateTime) into ICS UTC form, or "".
static string icsTimestamp(const string & text) {
    string date, time;
    if (!parseGraphDateTime(text, date, time)) return "";
    return date + "T" + time + "Z";
}

static void appendProperty(string & out, const string & nameAndParams, const string & value) {
    out += ICSFoldLine(nameAndParams + ":" + value);
}

// The first ':' of a content line ends its parameters, so it cannot appear in a quoted CN
// without confusing parsers that locate the value that way. DQUOTE is not allowed in a
// parameter value at all.
static string sanitizeParamValue(const string & value) {
    string out;
    for (unsigned char c : value) {
        if (c == '"' || c == ':' || c < 0x20) continue;
        out += (char)c;
    }
    return out;
}

static string trimmedAddress(const json & emailAddress) {
    string address = stringField(emailAddress, "address");
    string out;
    for (unsigned char c : address) {
        if (c > 0x20 && c != 0x7f) out += (char)c;
    }
    return out;
}

static string calAddressParams(const json & emailAddress) {
    string name = sanitizeParamValue(stringField(emailAddress, "name"));
    return name.empty() ? "" : ";CN=\"" + name + "\"";
}

static string partStat(const string & response) {
    // Graph reports the organizer's own entry as "organizer"; they have by definition accepted.
    if (response == "accepted" || response == "organizer") return "ACCEPTED";
    if (response == "declined") return "DECLINED";
    if (response == "tentativelyAccepted") return "TENTATIVE";
    return "NEEDS-ACTION";
}

bool GraphEventUnixBounds(const json & event, long long & rs, long long & re) {
    if (boolField(event, "isAllDay")) return false;
    const json * start = objectField(event, "start");
    if (start == nullptr) return false;
    long long startUnix = 0;
    if (!utcUnix(stringField(*start, "dateTime"), startUnix)) return false;
    long long endUnix = startUnix;
    if (const json * end = objectField(event, "end")) {
        if (!utcUnix(stringField(*end, "dateTime"), endUnix)) endUnix = startUnix;
    }
    rs = startUnix;
    re = endUnix;
    return true;
}

string GraphEventToICS(const json & event, vector<string> * warnings) {
    string graphId = stringField(event, "id");
    if (graphId.empty()) {
        throw invalid_argument("Graph event has no id");
    }
    // A control character would terminate the UID content line and inject further properties.
    for (unsigned char c : graphId) {
        if (c < 0x20) throw invalid_argument("Graph event id contains a control character");
    }

    bool allDay = boolField(event, "isAllDay");
    const json * start = objectField(event, "start");
    if (start == nullptr) {
        throw invalid_argument("Graph event has no start");
    }
    string dtStart = icsDateTime(*start, allDay, "start");
    string dtEnd;
    if (const json * end = objectField(event, "end")) {
        dtEnd = icsDateTime(*end, allDay, "end");
    }
    if (allDay && warnings != nullptr) {
        string startTime;
        string startDate;
        parseGraphDateTime(stringField(*start, "dateTime"), startDate, startTime);
        if (startTime != "000000") {
            warnings->push_back("all-day event " + graphId + " starts at " + startTime + ", not midnight; using its date");
        }
    }

    string modified = icsTimestamp(stringField(event, "lastModifiedDateTime"));
    if (modified.empty()) modified = icsTimestamp(stringField(event, "createdDateTime"));
    // DTSTAMP is mandatory; fall back to something deterministic so identical input yields
    // identical ICS (the sync skips saves whose ICS did not change).
    string dtStamp = modified.empty() ? (allDay ? dtStart + "T000000Z" : dtStart) : modified;

    string out;
    out += "BEGIN:VCALENDAR\r\n";
    out += "VERSION:2.0\r\n";
    out += "PRODID:-//Mailspring//Microsoft Graph Calendar//EN\r\n";
    out += "CALSCALE:GREGORIAN\r\n";
    out += "BEGIN:VEVENT\r\n";
    // Graph ids are base64url-like and need no TEXT escaping; keeping the UID verbatim keeps
    // it equal to GraphEventUID(), which the sync uses to find a stored event again.
    appendProperty(out, "UID", GraphEventUID(graphId));
    appendProperty(out, "DTSTAMP", dtStamp);
    if (!modified.empty()) {
        appendProperty(out, "LAST-MODIFIED", modified);
    }
    appendProperty(out, allDay ? "DTSTART;VALUE=DATE" : "DTSTART", dtStart);
    if (!dtEnd.empty()) {
        appendProperty(out, allDay ? "DTEND;VALUE=DATE" : "DTEND", dtEnd);
    }
    appendProperty(out, "SUMMARY", ICSEscapeText(stringField(event, "subject")));

    string description = stringField(event, "bodyPreview");
    if (!description.empty()) {
        appendProperty(out, "DESCRIPTION", ICSEscapeText(description));
    }
    if (const json * location = objectField(event, "location")) {
        string displayName = stringField(*location, "displayName");
        if (!displayName.empty()) {
            appendProperty(out, "LOCATION", ICSEscapeText(displayName));
        }
    }

    if (stringField(event, "showAs") == "free") {
        appendProperty(out, "TRANSP", "TRANSPARENT");
    }
    if (boolField(event, "isCancelled")) {
        appendProperty(out, "STATUS", "CANCELLED");
    }
    if (stringField(event, "sensitivity") == "private") {
        appendProperty(out, "CLASS", "PRIVATE");
    }

    if (const json * organizer = objectField(event, "organizer")) {
        if (const json * emailAddress = objectField(*organizer, "emailAddress")) {
            string address = trimmedAddress(*emailAddress);
            if (!address.empty()) {
                appendProperty(out, "ORGANIZER" + calAddressParams(*emailAddress), "mailto:" + address);
            }
        }
    }

    auto attendees = event.find("attendees");
    if (attendees != event.end() && attendees->is_array()) {
        for (const json & attendee : *attendees) {
            const json * emailAddress = objectField(attendee, "emailAddress");
            if (emailAddress == nullptr) continue;
            string address = trimmedAddress(*emailAddress);
            if (address.empty()) continue;

            string type = stringField(attendee, "type");
            string params;
            if (type == "resource") {
                params += ";CUTYPE=RESOURCE;ROLE=NON-PARTICIPANT";
            } else if (type == "optional") {
                params += ";ROLE=OPT-PARTICIPANT";
            } else {
                params += ";ROLE=REQ-PARTICIPANT";
            }
            // Graph names this `status` on attendees; accept `responseStatus` (the event-level name) too.
            const json * status = objectField(attendee, "status");
            if (status == nullptr) status = objectField(attendee, "responseStatus");
            params += ";PARTSTAT=" + partStat(status ? stringField(*status, "response") : "");
            params += calAddressParams(*emailAddress);
            appendProperty(out, "ATTENDEE" + params, "mailto:" + address);
        }
    }

    string joinUrl;
    if (const json * onlineMeeting = objectField(event, "onlineMeeting")) {
        joinUrl = stringField(*onlineMeeting, "joinUrl");
    }
    if (joinUrl.empty()) joinUrl = stringField(event, "onlineMeetingUrl");
    if (!joinUrl.empty()) {
        // URI values are not TEXT: no escaping, but line breaks would corrupt the stream.
        joinUrl.erase(remove_if(joinUrl.begin(), joinUrl.end(), [](unsigned char c) { return c < 0x20; }), joinUrl.end());
        appendProperty(out, "URL", joinUrl);
    }

    auto categories = event.find("categories");
    if (categories != event.end() && categories->is_array()) {
        string joined;
        for (const json & category : *categories) {
            if (!category.is_string()) continue;
            if (!joined.empty()) joined += ",";
            joined += ICSEscapeText(category.get<string>());
        }
        if (!joined.empty()) {
            appendProperty(out, "CATEGORIES", joined);
        }
    }

    string seriesMasterId = stringField(event, "seriesMasterId");
    if (!seriesMasterId.empty()) {
        appendProperty(out, "X-MS-SERIES-MASTER-ID", ICSEscapeText(seriesMasterId));
    }
    string iCalUId = stringField(event, "iCalUId");
    if (!iCalUId.empty()) {
        appendProperty(out, "X-MS-ICALUID", ICSEscapeText(iCalUId));
    }

    out += "END:VEVENT\r\n";
    out += "END:VCALENDAR\r\n";
    return out;
}

bool GraphIsSeriesMaster(const json & item) {
    return stringField(item, "type") == "seriesMaster";
}

bool GraphNeedsSeriesMerge(const json & item) {
    if (GraphIsSeriesMaster(item)) return false;
    if (stringField(item, "seriesMasterId").empty()) return false;
    return stringField(item, "subject").empty();
}

// Whether a field carries no information. Graph sends sparse occurrences without these keys,
// but placeholders such as a location with a blank displayName are treated the same way.
static bool isBlankField(const string & key, const json & value) {
    if (value.is_null()) return true;
    if (value.is_string()) return value.get<string>().empty();
    if (value.is_array() || value.is_object()) {
        if (value.empty()) return true;
        if (key == "location") return stringField(value, "displayName").empty();
        if (key == "body") return stringField(value, "content").empty();
    }
    return false;
}

json GraphMergeSeriesMaster(const json & occurrence, const json & master) {
    // Only descriptive fields come from the master; identity, timing and state stay the
    // occurrence's own (id, iCalUId, start, end, type, isCancelled, changeKey, ...).
    static const char * const inherited[] = {
        "subject", "bodyPreview", "body", "location", "locations", "organizer", "attendees",
        "categories", "showAs", "sensitivity", "importance", "isOnlineMeeting", "onlineMeeting",
        "onlineMeetingUrl", "isAllDay", "isReminderOn", "webLink",
    };
    json merged = occurrence;
    if (!occurrence.is_object() || !master.is_object()) return merged;
    for (const char * key : inherited) {
        auto source = master.find(key);
        if (source == master.end() || isBlankField(key, *source)) continue;
        auto own = merged.find(key);
        if (own != merged.end() && !isBlankField(key, *own)) continue;
        merged[key] = *source;
    }
    return merged;
}
