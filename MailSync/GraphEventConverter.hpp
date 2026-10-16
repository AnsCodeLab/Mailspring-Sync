//
//  GraphEventConverter.hpp
//  MailSync
//
//  Converts a Microsoft Graph event resource into an iCalendar document. Pure: no network,
//  no database.
//
//  Use of this file is subject to the terms and conditions defined
//  in 'LICENSE.md', which is part of the Mailspring-Sync package.
//

#ifndef GraphEventConverter_hpp
#define GraphEventConverter_hpp

#include <string>
#include <vector>
#include "json.hpp"

// Builds a VCALENDAR containing one VEVENT from a Graph event (as returned by
// calendarView/delta with `Prefer: outlook.timezone="UTC"`). Occurrences of a series are
// stored expanded, so the output never carries an RRULE.
//
// Throws std::invalid_argument when the event lacks an `id` (or the id holds a control
// character), has no usable `start`, or its times are not in UTC (the converter cannot apply
// time zone rules). `warnings`, when given, receives messages about data that was converted
// but looks wrong (e.g. an all-day event that does not start at midnight).
std::string GraphEventToICS(const nlohmann::json & event, std::vector<std::string> * warnings = nullptr);

// Event.rs/re for a timed Graph event: the true UTC instants of its start and end. The
// library's Date::toUnix() reads the same text as local time, which is wrong for `Z` times.
// Returns false for all-day events (their bounds are local midnights, which is what the app's
// day windows use, so the library's value is kept) and for unusable input.
bool GraphEventUnixBounds(const nlohmann::json & event, long long & rs, long long & re);

// calendarView/delta lists a recurring series' master (type "seriesMaster", full data, start =
// first occurrence) plus every occurrence as a sparse item without subject, body, location,
// organizer or attendees. The master is only a data source and is never stored as an event.
bool GraphIsSeriesMaster(const nlohmann::json & item);

// True for an occurrence/exception (has seriesMasterId) whose own subject is missing or empty.
bool GraphNeedsSeriesMerge(const nlohmann::json & item);

// Returns `occurrence` with its blank descriptive fields filled from `master`. Non-blank
// occurrence fields are kept; id, iCalUId, start, end, type, isCancelled, changeKey,
// lastModifiedDateTime and seriesMasterId are never taken from the master.
nlohmann::json GraphMergeSeriesMaster(const nlohmann::json & occurrence, const nlohmann::json & master);

// The ICS UID for a Graph event id. Every occurrence/exception of a series has its own Graph
// id, so the UID is derived from it rather than from the shared iCalUId.
std::string GraphEventUID(const std::string & graphEventId);

// RFC 5545 section 3.3.11 TEXT escaping (backslash, comma, semicolon, newlines).
std::string ICSEscapeText(const std::string & text);

// RFC 5545 section 3.1 folding: lines of at most 75 octets, continuation lines start with a
// space, multi-byte UTF-8 characters are never split. `line` must not contain CRLF; the
// result is CRLF terminated.
std::string ICSFoldLine(const std::string & line);

// Returns `ics` with every CR removed. Vendor/icalendarlib unfolds a continuation line
// without dropping the previous line's '\r', so a folded value (every long UID) parsed from
// CRLF text carries stray CRs; hand the library this LF-only form and store the CRLF original.
std::string ICSStripCR(const std::string & ics);

#endif /* GraphEventConverter_hpp */
