//
//  CalendarSyncWorker.hpp
//  MailSync
//
//  Use of this file is subject to the terms and conditions defined
//  in 'LICENSE.md', which is part of the Mailspring-Sync package.
//

#ifndef CalendarSyncWorker_hpp
#define CalendarSyncWorker_hpp

#include <string>

// Office 365 offers no CalDAV, so these accounts sync calendars through Microsoft Graph.
inline bool ProviderUsesGraphCalendar(const std::string & provider) {
    return provider == "office365" || provider == "outlook";
}

// One pass of the calendar (and, for CalDAV/CardDAV, contacts) sync for a single account.
// runCalContactsSyncWorker() in main.cpp owns the scheduling and error policy, so the
// implementations only throw SyncException and never loop or sleep between passes.
class CalendarSyncWorker {
public:
    virtual ~CalendarSyncWorker() = default;
    // Everything this worker syncs: calendars plus, where the provider has any, contacts.
    virtual void run() = 0;
    // Calendars only; used on the passes where contacts are not due.
    virtual void runCalendars() = 0;
};

#endif /* CalendarSyncWorker_hpp */
