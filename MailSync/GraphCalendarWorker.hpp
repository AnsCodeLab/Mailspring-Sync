//
//  GraphCalendarWorker.hpp
//  MailSync
//
//  Read-only calendar sync for Microsoft 365 / Outlook accounts through Microsoft Graph.
//  Calendars and events land in the same Calendar and Event tables the CalDAV sync fills.
//
//  Use of this file is subject to the terms and conditions defined
//  in 'LICENSE.md', which is part of the Mailspring-Sync package.
//

#ifndef GraphCalendarWorker_hpp
#define GraphCalendarWorker_hpp

#include <memory>
#include <set>
#include <string>
#include <vector>

#include "Account.hpp"
#include "Calendar.hpp"
#include "CalendarSyncWorker.hpp"
#include "GraphDeltaState.hpp"
#include "MailStore.hpp"
#include "json.hpp"
#include "spdlog/spdlog.h"

using namespace std;

class GraphCalendarWorker : public CalendarSyncWorker {
    // Like every MailStore user this is bound to the thread that constructed the worker.
    unique_ptr<MailStore> store;
    shared_ptr<Account> account;
    shared_ptr<spdlog::logger> logger;

    json graphGet(const string & url);

    void syncCalendarList();
    void syncCalendar(shared_ptr<Calendar> calendar);
    void runDelta(shared_ptr<Calendar> calendar, const string & startURL, const GraphDeltaState & window, bool full);
    void applyDeltaPage(shared_ptr<Calendar> calendar, const json & items, set<string> * seenEventIds);
    void removeEvents(vector<string> eventIds);
    void removeEventsNotIn(const string & calendarId, const set<string> & keepEventIds);
    void removeCalendar(shared_ptr<Calendar> calendar);

public:
    GraphCalendarWorker(shared_ptr<Account> account);

    void run() override;
};

#endif /* GraphCalendarWorker_hpp */
