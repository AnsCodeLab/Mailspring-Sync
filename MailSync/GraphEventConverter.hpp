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
#include "json.hpp"

// Builds a VCALENDAR containing one VEVENT from a Graph event (as returned by
// calendarView/delta with `Prefer: outlook.timezone="UTC"`). Occurrences of a series are
// stored expanded, so the output never carries an RRULE.
//
// Throws std::invalid_argument when the event lacks an `id`, has no usable `start`, or its
// times are not in UTC (the converter cannot apply time zone rules).
std::string GraphEventToICS(const nlohmann::json & event);

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
