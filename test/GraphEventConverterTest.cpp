// Unit tests for the Microsoft Graph calendar sync helpers that need no network or database:
// GraphEventConverter, GraphDeltaState and the OAuth token cache. Built only with
// -DMAILSYNC_BUILD_TESTS=ON; plain checks, no test framework.

#include <cstdio>
#include <cstdlib>
#include <ctime>
#include <string>
#include <vector>

#include "GraphDeltaState.hpp"
#include "GraphEventConverter.hpp"
#include "XOAuth2TokenCache.hpp"
#include "icalendar.h"
#include "json.hpp"

using namespace std;
using nlohmann::json;

static int failures = 0;
static int checks = 0;

#define CHECK(cond)                                                                   \
    do {                                                                              \
        checks++;                                                                     \
        if (!(cond)) {                                                                \
            failures++;                                                               \
            printf("  FAIL %s:%d: %s\n", __FILE__, __LINE__, #cond);                  \
        }                                                                             \
    } while (0)

#define CHECK_EQ(actual, expected)                                                    \
    do {                                                                              \
        checks++;                                                                     \
        string a_ = (actual);                                                         \
        string e_ = (expected);                                                       \
        if (a_ != e_) {                                                               \
            failures++;                                                               \
            printf("  FAIL %s:%d: %s\n    actual:   [%s]\n    expected: [%s]\n",      \
                   __FILE__, __LINE__, #actual, a_.c_str(), e_.c_str());              \
        }                                                                             \
    } while (0)

static bool contains(const string & haystack, const string & needle) {
    return haystack.find(needle) != string::npos;
}

static bool hasLine(const string & ics, const string & line) {
    return contains(ics, "\r\n" + line + "\r\n") || ics.find(line + "\r\n") == 0;
}

static string unfold(const string & ics) {
    string out;
    for (size_t i = 0; i < ics.size(); i++) {
        if (ics[i] == '\r' && ics.compare(i, 3, "\r\n ") == 0) {
            i += 2;
            continue;
        }
        out += ics[i];
    }
    return out;
}

// Vendor/icalendarlib is what the engine parses stored ICS with (see ICSStripCR).
static shared_ptr<ICalendar> parse(const string & ics) {
    return make_shared<ICalendar>(ICSStripCR(ics));
}

static json baseEvent() {
    return json::parse(R"({
        "id": "AAMkAGI2TG93AAA=",
        "iCalUId": "040000008200E00074C5B7101A82E008",
        "changeKey": "DwAAABYAAAB",
        "subject": "Weekly sync",
        "bodyPreview": "Agenda",
        "isAllDay": false,
        "isCancelled": false,
        "sensitivity": "normal",
        "showAs": "busy",
        "start": {"dateTime": "2025-01-15T10:30:00.0000000", "timeZone": "UTC"},
        "end": {"dateTime": "2025-01-15T11:00:00.0000000", "timeZone": "UTC"},
        "lastModifiedDateTime": "2025-01-10T08:15:30.1234567Z",
        "location": {"displayName": "Room 4"},
        "organizer": {"emailAddress": {"name": "Alice Org", "address": "alice@example.com"}},
        "attendees": []
    })");
}

static void testTimedEvent() {
    printf("timed event\n");
    string ics = GraphEventToICS(baseEvent());
    CHECK(ics.find("BEGIN:VCALENDAR\r\n") == 0);
    CHECK(contains(ics, "END:VCALENDAR\r\n"));
    CHECK(contains(ics, "PRODID:"));
    CHECK(hasLine(ics, "DTSTART:20250115T103000Z"));
    CHECK(hasLine(ics, "DTEND:20250115T110000Z"));
    CHECK(hasLine(ics, "DTSTAMP:20250110T081530Z"));
    CHECK(hasLine(ics, "LAST-MODIFIED:20250110T081530Z"));
    CHECK(!contains(ics, "RRULE"));

    auto cal = parse(ics);
    CHECK(cal->Events.size() == 1);
    ICalendarEvent * e = cal->Events.front();
    CHECK_EQ(e->UID, "AAMkAGI2TG93AAA=@graph.microsoft.com");
    CHECK_EQ(e->Summary, "Weekly sync");
    CHECK_EQ(e->Description, "Agenda");
    CHECK_EQ(e->Location, "Room 4");
    CHECK_EQ(string(e->DtStart), "20250115T103000");
    CHECK_EQ(string(e->DtEnd), "20250115T110000");
    CHECK(e->DtStart.WithTime);
    CHECK_EQ(e->Organizer, "Alice Org <alice@example.com>");
    CHECK(e->Status.empty());
    CHECK(e->RRule.IsEmpty());
}

static void testAllDayEvent() {
    printf("all-day event\n");
    json ev = baseEvent();
    ev["isAllDay"] = true;
    ev["start"] = {{"dateTime", "2025-01-16T00:00:00.0000000"}, {"timeZone", "UTC"}};
    ev["end"] = {{"dateTime", "2025-01-17T00:00:00.0000000"}, {"timeZone", "UTC"}};
    string ics = GraphEventToICS(ev);
    CHECK(hasLine(ics, "DTSTART;VALUE=DATE:20250116"));
    CHECK(hasLine(ics, "DTEND;VALUE=DATE:20250117"));

    auto cal = parse(ics);
    CHECK(cal->Events.size() == 1);
    ICalendarEvent * e = cal->Events.front();
    CHECK_EQ(string(e->DtStart), "20250116");
    CHECK_EQ(string(e->DtEnd), "20250117");
    CHECK(!e->DtStart.WithTime);
}

static void testFreePrivateCancelled() {
    printf("free / private / cancelled\n");
    json ev = baseEvent();
    ev["showAs"] = "free";
    ev["sensitivity"] = "private";
    ev["isCancelled"] = true;
    string ics = GraphEventToICS(ev);
    CHECK(hasLine(ics, "TRANSP:TRANSPARENT"));
    CHECK(hasLine(ics, "CLASS:PRIVATE"));
    CHECK(hasLine(ics, "STATUS:CANCELLED"));
    CHECK_EQ(parse(ics)->Events.front()->Status, "CANCELLED");

    string plain = GraphEventToICS(baseEvent());
    CHECK(!contains(plain, "TRANSP"));
    CHECK(!contains(plain, "CLASS:"));
    CHECK(!contains(plain, "STATUS:"));
}

static void testAttendeesAndOrganizer() {
    printf("attendees / organizer\n");
    json ev = baseEvent();
    ev["attendees"] = json::array({
        {{"type", "required"}, {"status", {{"response", "accepted"}}}, {"emailAddress", {{"name", "Bob"}, {"address", "bob@example.com"}}}},
        {{"type", "optional"}, {"status", {{"response", "declined"}}}, {"emailAddress", {{"name", "Carol"}, {"address", "carol@example.com"}}}},
        {{"type", "required"}, {"status", {{"response", "tentativelyAccepted"}}}, {"emailAddress", {{"name", "Dan"}, {"address", "dan@example.com"}}}},
        {{"type", "required"}, {"status", {{"response", "none"}}}, {"emailAddress", {{"name", "Eve: \"E\""}, {"address", "eve@example.com"}}}},
        {{"type", "resource"}, {"responseStatus", {{"response", "accepted"}}}, {"emailAddress", {{"name", "Room 4"}, {"address", "room4@example.com"}}}},
        {{"type", "required"}, {"emailAddress", {{"name", "No address"}, {"address", ""}}}},
    });
    string ics = GraphEventToICS(ev);
    string flat = unfold(ics);
    CHECK(contains(flat, "ATTENDEE;ROLE=REQ-PARTICIPANT;PARTSTAT=ACCEPTED;CN=\"Bob\":mailto:bob@example.com\r\n"));
    CHECK(contains(flat, "ATTENDEE;ROLE=OPT-PARTICIPANT;PARTSTAT=DECLINED;CN=\"Carol\":mailto:carol@example.com\r\n"));
    CHECK(contains(flat, "PARTSTAT=TENTATIVE;CN=\"Dan\":mailto:dan@example.com"));
    CHECK(contains(flat, "PARTSTAT=NEEDS-ACTION;CN=\"Eve E\":mailto:eve@example.com"));
    CHECK(contains(flat, "ATTENDEE;CUTYPE=RESOURCE;ROLE=NON-PARTICIPANT;PARTSTAT=ACCEPTED;CN=\"Room 4\":mailto:room4@example.com"));
    CHECK(!contains(flat, "No address"));
    CHECK(contains(flat, "ORGANIZER;CN=\"Alice Org\":mailto:alice@example.com\r\n"));

    auto cal = parse(ics);
    ICalendarEvent * e = cal->Events.front();
    CHECK_EQ(e->Organizer, "Alice Org <alice@example.com>");
    vector<string> attendees(e->Attendees.begin(), e->Attendees.end());
    CHECK(attendees.size() == 5);
    CHECK_EQ(attendees[0], "Bob <bob@example.com>");
    CHECK_EQ(attendees[1], "Carol <carol@example.com>");
    CHECK_EQ(attendees[3], "Eve E <eve@example.com>");
}

static void testOnlineMeetingAndCategories() {
    printf("url / categories\n");
    json ev = baseEvent();
    ev["onlineMeeting"] = {{"joinUrl", "https://teams.microsoft.com/l/meetup-join/abc"}};
    ev["categories"] = json::array({"Red, team", "Blue"});
    string ics = GraphEventToICS(ev);
    CHECK(hasLine(ics, "URL:https://teams.microsoft.com/l/meetup-join/abc"));
    CHECK(hasLine(ics, "CATEGORIES:Red\\, team,Blue"));

    json legacy = baseEvent();
    legacy["onlineMeetingUrl"] = "https://example.com/join";
    CHECK(hasLine(GraphEventToICS(legacy), "URL:https://example.com/join"));
}

static void testEscaping() {
    printf("text escaping\n");
    CHECK_EQ(ICSEscapeText("a,b;c\\d"), "a\\,b\\;c\\\\d");
    CHECK_EQ(ICSEscapeText("line1\nline2\r\nline3\rline4"), "line1\\nline2\\nline3\\nline4");
    CHECK_EQ(ICSEscapeText("Caf\xC3\xA9 \xE6\x97\xA5\xE6\x9C\xAC \xF0\x9F\x93\x85"), "Caf\xC3\xA9 \xE6\x97\xA5\xE6\x9C\xAC \xF0\x9F\x93\x85");
    CHECK_EQ(ICSEscapeText(string("x\x01y\tz")), "xy\tz");

    json ev = baseEvent();
    ev["subject"] = "Q1, plan; review \xE6\x97\xA5\xE6\x9C\xAC";
    ev["location"] = {{"displayName", "Hall A, Floor 2; East\nWing \xC3\xA9"}};
    string ics = GraphEventToICS(ev);
    CHECK(contains(ics, "SUMMARY:Q1\\, plan\\; review \xE6\x97\xA5\xE6\x9C\xAC\r\n"));
    // The library unescapes LOCATION on parse.
    CHECK_EQ(parse(ics)->Events.front()->Location, "Hall A, Floor 2; East\nWing \xC3\xA9");
    CHECK(!contains(unfold(ics), "\r\nWing"));
}

static void testFolding() {
    printf("line folding\n");
    CHECK_EQ(ICSFoldLine("SHORT:value"), "SHORT:value\r\n");
    CHECK_EQ(ICSFoldLine(""), "\r\n");

    string exact = "X:" + string(73, 'a');  // exactly 75 octets: not folded
    CHECK_EQ(ICSFoldLine(exact), exact + "\r\n");
    string over = "X:" + string(74, 'a');    // 76 octets: folds
    CHECK(contains(ICSFoldLine(over), "\r\n a\r\n"));

    // 2-octet characters straddling every possible fold position
    for (int pad = 0; pad < 4; pad++) {
        string line = "SUMMARY:" + string(pad, 'x');
        for (int i = 0; i < 120; i++) line += "\xC3\xA9";
        for (int i = 0; i < 40; i++) line += "\xF0\x9F\x93\x85";
        string folded = ICSFoldLine(line);

        size_t pos = 0;
        bool allShort = true, noSplitChar = true;
        while (pos < folded.size()) {
            size_t eol = folded.find("\r\n", pos);
            string physical = folded.substr(pos, eol - pos);
            if (physical.size() > 75) allShort = false;
            // a physical line must neither end on a lead byte nor start (after the space) on a continuation byte
            size_t bodyStart = pos == 0 ? 0 : 1;
            if (physical.size() > bodyStart && (static_cast<unsigned char>(physical[bodyStart]) & 0xC0) == 0x80) noSplitChar = false;
            pos = eol + 2;
        }
        CHECK(allShort);
        CHECK(noSplitChar);
        CHECK_EQ(unfold(folded), line + "\r\n");
    }

    json ev = baseEvent();
    string subject;
    for (int i = 0; i < 90; i++) subject += "Ab\xC3\xA9 ";
    ev["subject"] = subject;
    string ics = GraphEventToICS(ev);
    size_t pos = 0;
    while (pos < ics.size()) {
        size_t eol = ics.find("\r\n", pos);
        CHECK(eol != string::npos);
        CHECK(eol - pos <= 75);
        pos = eol + 2;
    }
    CHECK_EQ(parse(ics)->Events.front()->Summary, subject);
}

static void testSeriesOccurrence() {
    printf("series occurrence / uid\n");
    json a = baseEvent();
    a["id"] = "AAMkOccurrenceA=";
    a["type"] = "occurrence";
    a["seriesMasterId"] = "AAMkMaster=";
    json b = a;
    b["id"] = "AAMkOccurrenceB=";
    b["start"]["dateTime"] = "2025-01-22T10:30:00.0000000";
    b["end"]["dateTime"] = "2025-01-22T11:00:00.0000000";

    string icsA = GraphEventToICS(a);
    string icsB = GraphEventToICS(b);
    CHECK(hasLine(icsA, "X-MS-SERIES-MASTER-ID:AAMkMaster="));
    CHECK(hasLine(icsA, "X-MS-ICALUID:040000008200E00074C5B7101A82E008"));
    CHECK(!contains(icsA, "RRULE"));
    CHECK(!contains(icsA, "RECURRENCE-ID"));

    string uidA = parse(icsA)->Events.front()->UID;
    string uidB = parse(icsB)->Events.front()->UID;
    CHECK_EQ(uidA, GraphEventUID("AAMkOccurrenceA="));
    CHECK_EQ(uidB, GraphEventUID("AAMkOccurrenceB="));
    CHECK(uidA != uidB);  // same iCalUId, distinct events

    // A realistic ~150 character Graph id folds the UID line; the parsed UID must stay intact.
    string longId = "AAMkAGI2THVSAAA=" + string(130, 'Z') + "-_q=";
    json c = baseEvent();
    c["id"] = longId;
    string icsC = GraphEventToICS(c);
    CHECK(contains(icsC, "\r\n "));  // the UID line was folded
    CHECK(contains(icsC, "UID:" + longId.substr(0, 20)));
    CHECK_EQ(parse(icsC)->Events.front()->UID, GraphEventUID(longId));
}

static void testInvalidInput() {
    printf("invalid input\n");
    auto throwsInvalid = [](const json & ev) {
        try {
            GraphEventToICS(ev);
        } catch (invalid_argument &) {
            return true;
        }
        return false;
    };
    json noId = baseEvent();
    noId.erase("id");
    CHECK(throwsInvalid(noId));

    json noStart = baseEvent();
    noStart.erase("start");
    CHECK(throwsInvalid(noStart));

    json badStart = baseEvent();
    badStart["start"] = {{"dateTime", "not a date"}, {"timeZone", "UTC"}};
    CHECK(throwsInvalid(badStart));

    json zoned = baseEvent();
    zoned["start"]["timeZone"] = "Pacific Standard Time";
    CHECK(throwsInvalid(zoned));

    CHECK(throwsInvalid(json::array()));

    // null / odd typed optional fields are ignored rather than fatal
    json sparse = baseEvent();
    sparse["location"] = nullptr;
    sparse["organizer"] = nullptr;
    sparse["attendees"] = nullptr;
    sparse["subject"] = nullptr;
    sparse["isCancelled"] = nullptr;
    string ics = GraphEventToICS(sparse);
    CHECK(parse(ics)->Events.size() == 1);
}

static void testDeltaBookkeeping() {
    printf("delta bookkeeping\n");
    GraphDeltaState state;
    state.deltaLink = "https://graph.microsoft.com/v1.0/me/calendars/ID/calendarView/delta?$deltatoken=abc";
    state.windowStart = 1700000000;
    state.windowEnd = 1760000000;
    string token = GraphDeltaStateToToken(state);

    GraphDeltaState parsed;
    CHECK(GraphDeltaStateFromToken(token, parsed));
    CHECK_EQ(parsed.deltaLink, state.deltaLink);
    CHECK(parsed.windowStart == state.windowStart);
    CHECK(parsed.windowEnd == state.windowEnd);

    GraphDeltaState untouched;
    untouched.windowStart = 42;
    CHECK(!GraphDeltaStateFromToken("", untouched));
    CHECK(!GraphDeltaStateFromToken("http://sync.example.com/token/1234", untouched));
    CHECK(!GraphDeltaStateFromToken("{", untouched));
    CHECK(!GraphDeltaStateFromToken("{}", untouched));
    CHECK(!GraphDeltaStateFromToken("{\"deltaLink\":\"\",\"windowStart\":1,\"windowEnd\":2}", untouched));
    CHECK(!GraphDeltaStateFromToken("{\"deltaLink\":\"x\",\"windowStart\":\"1\",\"windowEnd\":2}", untouched));
    CHECK(untouched.windowStart == 42);

    time_t day = 24 * 60 * 60;
    CHECK(GraphNeedsFullResync("", 1700000000));
    CHECK(GraphNeedsFullResync("http://sync.example.com/token/1234", 1700000000));
    CHECK(!GraphNeedsFullResync(token, 1700000000));                       // same window
    CHECK(!GraphNeedsFullResync(token, 1700000000 + 3 * day));            // slight drift
    CHECK(!GraphNeedsFullResync(token, 1700000000 + 7 * day));            // exactly 7 days: still incremental
    CHECK(GraphNeedsFullResync(token, 1700000000 + 7 * day + 1));         // more than 7 days
    CHECK(!GraphNeedsFullResync(token, 1700000000 - 30 * day));           // desired start earlier: keep

    CHECK_EQ(GraphFormatISO8601UTC(0), "1970-01-01T00:00:00Z");
    CHECK_EQ(GraphFormatISO8601UTC(1736937000), "2025-01-15T10:30:00Z");
    CHECK_EQ(GraphUrlEncode("AAMk=/+x-_.~"), "AAMk%3D%2F%2Bx-_.~");
    CHECK_EQ(GraphDeltaStartURL("AA=", 1736937000, 1736937000 + 86400),
             "https://graph.microsoft.com/v1.0/me/calendars/AA%3D/calendarView/delta"
             "?startDateTime=2025-01-15T10%3A30%3A00Z&endDateTime=2025-01-16T10%3A30%3A00Z");

    CHECK(GraphBoundedRetryAfterSec("") == 5);
    CHECK(GraphBoundedRetryAfterSec("3") == 3);
    CHECK(GraphBoundedRetryAfterSec("0") == 1);
    CHECK(GraphBoundedRetryAfterSec("9999") == 60);
    CHECK(GraphBoundedRetryAfterSec("99999999999999999999") == 60);
    CHECK(GraphBoundedRetryAfterSec("Fri, 31 Dec 2038 23:59:59 GMT") == 5);
    CHECK(GraphBoundedRetryAfterSec("-4") == 5);

    CHECK_EQ(GraphCalendarColor("#ff8800", "lightBlue"), "#ff8800");
    CHECK_EQ(GraphCalendarColor("", "lightGreen"), "#7fba00");
    CHECK_EQ(GraphCalendarColor("", "auto"), "#4285f4");
    CHECK_EQ(GraphCalendarColor("", ""), "#4285f4");
}

static void testTokenCache() {
    printf("token cache\n");
    CHECK(XOAuth2CacheKey("acct", OAuthResource::Mail) != XOAuth2CacheKey("acct", OAuthResource::Graph));
    CHECK(XOAuth2CacheKey("acct", OAuthResource::Mail) != XOAuth2CacheKey("acct2", OAuthResource::Mail));

    XOAuth2TokenCache cache;
    time_t now = 1000000;
    XOAuth2Parts out;
    CHECK(!cache.lookup("acct", OAuthResource::Mail, now, 60, out));

    cache.store("acct", OAuthResource::Mail, {"u", "mail-token", now + 3600});
    cache.store("acct", OAuthResource::Graph, {"u", "graph-token", now + 3600});
    CHECK(cache.lookup("acct", OAuthResource::Mail, now, 60, out));
    CHECK_EQ(out.accessToken, "mail-token");
    CHECK(cache.lookup("acct", OAuthResource::Graph, now, 60, out));
    CHECK_EQ(out.accessToken, "graph-token");

    // refreshing the graph token must not evict the mail token (and vice versa)
    cache.store("acct", OAuthResource::Graph, {"u", "graph-token-2", now + 7200});
    CHECK(cache.lookup("acct", OAuthResource::Mail, now, 60, out));
    CHECK_EQ(out.accessToken, "mail-token");
    CHECK(cache.lookup("acct", OAuthResource::Graph, now, 60, out));
    CHECK_EQ(out.accessToken, "graph-token-2");
    CHECK(!cache.lookup("other", OAuthResource::Graph, now, 60, out));

    // 60 s buffer: a token expiring within the buffer counts as expired, one just beyond does not
    cache.store("acct", OAuthResource::Mail, {"u", "short", now + 60});
    CHECK(!cache.lookup("acct", OAuthResource::Mail, now, 60, out));
    cache.store("acct", OAuthResource::Mail, {"u", "short", now + 61});
    CHECK(cache.lookup("acct", OAuthResource::Mail, now, 60, out));
    CHECK(cache.lookup("acct", OAuthResource::Graph, now, 60, out));
}

static void setZone(const char * zone) {
    setenv("TZ", zone, 1);
    tzset();
}

// Local-midnight instant of a calendar day in the current process time zone.
static long long localMidnight(int year, int month, int day) {
    tm parts {};
    parts.tm_year = year - 1900;
    parts.tm_mon = month - 1;
    parts.tm_mday = day;
    parts.tm_isdst = -1;
    return (long long)mktime(&parts);
}

static void testUtcBoundsUnderTimeZones() {
    printf("rs/re under non-UTC process time zones\n");
    json timed = baseEvent();
    timed["start"] = {{"dateTime", "2025-01-15T09:00:00.0000000"}, {"timeZone", "UTC"}};
    timed["end"] = {{"dateTime", "2025-01-15T10:00:00.0000000"}, {"timeZone", "UTC"}};
    const long long trueStart = 1736931600;  // 2025-01-15T09:00:00Z
    const long long trueEnd = 1736935200;    // 2025-01-15T10:00:00Z

    json allDay = baseEvent();
    allDay["isAllDay"] = true;
    allDay["start"] = {{"dateTime", "2025-01-16T00:00:00.0000000"}, {"timeZone", "UTC"}};
    allDay["end"] = {{"dateTime", "2025-01-17T00:00:00.0000000"}, {"timeZone", "UTC"}};

    const char * zones[] = {"UTC", "Asia/Ho_Chi_Minh", "America/Los_Angeles"};
    for (const char * zone : zones) {
        setZone(zone);
        printf("  TZ=%s\n", zone);

        long long rs = 0, re = 0;
        CHECK(GraphEventUnixBounds(timed, rs, re));
        CHECK(rs == trueStart);
        CHECK(re == trueEnd);

        // The library alone reads the same text as local time, which is the bug the helper fixes.
        auto cal = parse(GraphEventToICS(timed));
        long long libraryStart = cal->Events.front()->DtStart.toUnix();
        if (string(zone) == "UTC") {
            CHECK(libraryStart == trueStart);
        } else {
            CHECK(libraryStart != trueStart);
        }

        // All-day events keep local-midnight bounds from the library and get no UTC override.
        long long ars = -1, are = -1;
        CHECK(!GraphEventUnixBounds(allDay, ars, are));
        CHECK(ars == -1 && are == -1);
        auto dayCal = parse(GraphEventToICS(allDay));
        CHECK(dayCal->Events.front()->DtStart.toUnix() == localMidnight(2025, 1, 16));
        CHECK(dayCal->Events.front()->DtEnd.toUnix() == localMidnight(2025, 1, 17));
    }
    setZone("UTC");

    // different zones really do give different local midnights (the check above is not vacuous)
    setZone("Asia/Ho_Chi_Minh");
    long long hcm = localMidnight(2025, 1, 16);
    setZone("America/Los_Angeles");
    long long la = localMidnight(2025, 1, 16);
    CHECK(hcm != la);
    setZone("UTC");

    // missing end: re falls back to rs; unusable start yields no override
    json noEnd = timed;
    noEnd.erase("end");
    long long rs = 0, re = 0;
    CHECK(GraphEventUnixBounds(noEnd, rs, re));
    CHECK(rs == trueStart && re == trueStart);
    json bad = timed;
    bad["start"] = {{"dateTime", "garbage"}, {"timeZone", "UTC"}};
    CHECK(!GraphEventUnixBounds(bad, rs, re));
}

static void testZoneNames() {
    printf("time zone names\n");
    auto accepts = [](const char * zone) {
        json ev = baseEvent();
        ev["start"]["timeZone"] = zone;
        ev["end"]["timeZone"] = zone;
        try {
            GraphEventToICS(ev);
        } catch (invalid_argument &) {
            return false;
        }
        return true;
    };
    CHECK(accepts("UTC"));
    CHECK(accepts("utc"));
    CHECK(accepts("Etc/UTC"));
    CHECK(accepts("ETC/utc"));
    CHECK(accepts("tzone://Microsoft/Utc"));
    CHECK(accepts("Greenwich Standard Time"));
    CHECK(accepts("greenwich standard time"));
    CHECK(!accepts("GMT Standard Time"));       // London: observes DST
    CHECK(!accepts("UTC+01:00"));
    CHECK(!accepts("Etc/UTC-1"));
    CHECK(!accepts("Pacific Standard Time"));
    CHECK(!accepts("W. Europe Standard Time"));
}

static void testConverterEdgeCases() {
    printf("converter edge cases\n");

    // missing end: no DTEND line, the library reports an empty end
    json noEnd = baseEvent();
    noEnd.erase("end");
    string ics = GraphEventToICS(noEnd);
    CHECK(!contains(ics, "DTEND"));
    CHECK(parse(ics)->Events.front()->DtEnd.IsEmpty());
    CHECK(hasLine(ics, "DTSTART:20250115T103000Z"));

    // empty subject stays a (blank) SUMMARY and parses
    json empty = baseEvent();
    empty["subject"] = "";
    ics = GraphEventToICS(empty);
    CHECK(hasLine(ics, "SUMMARY:"));
    CHECK_EQ(parse(ics)->Events.front()->Summary, "");

    // DTSTAMP fallback is deterministic: start for timed, midnight of the start date for all-day
    json noStamp = baseEvent();
    noStamp.erase("lastModifiedDateTime");
    ics = GraphEventToICS(noStamp);
    CHECK(hasLine(ics, "DTSTAMP:20250115T103000Z"));
    CHECK(!contains(ics, "LAST-MODIFIED"));
    CHECK_EQ(ics, GraphEventToICS(noStamp));
    json created = noStamp;
    created["createdDateTime"] = "2025-01-01T01:02:03.0000000Z";
    CHECK(hasLine(GraphEventToICS(created), "DTSTAMP:20250101T010203Z"));
    json dayNoStamp = noStamp;
    dayNoStamp["isAllDay"] = true;
    dayNoStamp["start"] = {{"dateTime", "2025-01-16T00:00:00.0000000"}, {"timeZone", "UTC"}};
    dayNoStamp["end"] = {{"dateTime", "2025-01-17T00:00:00.0000000"}, {"timeZone", "UTC"}};
    CHECK(hasLine(GraphEventToICS(dayNoStamp), "DTSTAMP:20250116T000000Z"));

    // all-day event that does not start at midnight: date part used, warning reported
    json odd = baseEvent();
    odd["isAllDay"] = true;
    odd["start"] = {{"dateTime", "2025-01-16T08:00:00.0000000"}, {"timeZone", "UTC"}};
    odd["end"] = {{"dateTime", "2025-01-17T08:00:00.0000000"}, {"timeZone", "UTC"}};
    vector<string> warnings;
    ics = GraphEventToICS(odd, &warnings);
    CHECK(hasLine(ics, "DTSTART;VALUE=DATE:20250116"));
    CHECK(warnings.size() == 1);
    CHECK(!warnings.empty() && contains(warnings[0], "080000"));
    warnings.clear();
    json midnight = odd;
    midnight["start"]["dateTime"] = "2025-01-16T00:00:00.0000000";
    GraphEventToICS(midnight, &warnings);
    CHECK(warnings.empty());

    // organizer's own response counts as accepted; other attendee responses unchanged
    json org = baseEvent();
    org["attendees"] = json::array({
        {{"type", "required"}, {"status", {{"response", "organizer"}}}, {"emailAddress", {{"name", "Alice Org"}, {"address", "alice@example.com"}}}},
        {{"type", "required"}, {"status", {{"response", "notResponded"}}}, {"emailAddress", {{"name", "Bob"}, {"address", "bob@example.com"}}}},
    });
    string flat = unfold(GraphEventToICS(org));
    CHECK(contains(flat, "PARTSTAT=ACCEPTED;CN=\"Alice Org\":mailto:alice@example.com"));
    CHECK(contains(flat, "PARTSTAT=NEEDS-ACTION;CN=\"Bob\":mailto:bob@example.com"));

    // a control character in the id would break out of the UID content line
    json ctl = baseEvent();
    ctl["id"] = string("AAMk\r\nATTENDEE:mailto:evil@example.com");
    bool threw = false;
    try {
        GraphEventToICS(ctl);
    } catch (invalid_argument &) {
        threw = true;
    }
    CHECK(threw);
    ctl["id"] = string("AAMk\tTab");
    threw = false;
    try {
        GraphEventToICS(ctl);
    } catch (invalid_argument &) {
        threw = true;
    }
    CHECK(threw);
}

static void testUrlTrustAndStaleToken() {
    printf("trusted urls / stale token / unusable pass\n");
    CHECK(GraphIsTrustedURL("https://graph.microsoft.com/v1.0/me/calendars"));
    CHECK(GraphIsTrustedURL("https://graph.microsoft.com/v1.0/me/calendars/X/calendarView/delta?$deltatoken=a"));
    CHECK(!GraphIsTrustedURL(""));
    CHECK(!GraphIsTrustedURL("http://graph.microsoft.com/v1.0/me"));
    CHECK(!GraphIsTrustedURL("https://graph.microsoft.com"));
    CHECK(!GraphIsTrustedURL("https://graph.microsoft.com.evil.example/v1.0/me"));
    CHECK(!GraphIsTrustedURL("https://graph.microsoft.com@evil.example/v1.0/me"));
    CHECK(!GraphIsTrustedURL("https://evil.example/https://graph.microsoft.com/"));
    CHECK(!GraphIsTrustedURL("HTTPS://GRAPH.MICROSOFT.COM/v1.0"));  // strict: Graph returns lower case
    CHECK(!GraphIsTrustedURL("https://graph.microsoft.de/v1.0/me"));

    CHECK(GraphStatusMeansStaleDeltaToken("Invalid Response Code: 410"));
    CHECK(GraphStatusMeansStaleDeltaToken("Invalid Response Code: 400"));
    CHECK(!GraphStatusMeansStaleDeltaToken("Invalid Response Code: 404"));
    CHECK(!GraphStatusMeansStaleDeltaToken("Invalid Response Code: 500"));
    CHECK(!GraphStatusMeansStaleDeltaToken("Invalid Response Code: 401"));

    CHECK(!GraphPassUnusable(0, 0));   // nothing received: nothing wrong
    CHECK(!GraphPassUnusable(5, 0));
    CHECK(!GraphPassUnusable(5, 4));
    CHECK(GraphPassUnusable(5, 5));
    CHECK(GraphPassUnusable(1, 1));
}

int main() {
    testTimedEvent();
    testAllDayEvent();
    testFreePrivateCancelled();
    testAttendeesAndOrganizer();
    testOnlineMeetingAndCategories();
    testEscaping();
    testFolding();
    testSeriesOccurrence();
    testInvalidInput();
    testDeltaBookkeeping();
    testTokenCache();
    testUtcBoundsUnderTimeZones();
    testZoneNames();
    testConverterEdgeCases();
    testUrlTrustAndStaleToken();

    printf("%d checks, %d failure(s)\n", checks, failures);
    return failures == 0 ? 0 : 1;
}
