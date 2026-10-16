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

#include <map>
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
    // Bound to the thread that constructed the worker, like every MailStore; never freed
    // (see the constructor).
    MailStore * store;
    shared_ptr<Account> account;
    shared_ptr<spdlog::logger> logger;

    json graphGet(const string & url);

    // State of the delta pass in flight (a worker runs one pass at a time): series masters by
    // Graph id (null = unavailable), whether any master was delivered, and how many items of
    // the pass were masters, which are not events and so do not count as received events.
    map<string, json> masterCache;
    bool sawSeriesMaster = false;
    size_t passMasterItems = 0;

    void syncCalendarList();
    void syncCalendar(shared_ptr<Calendar> calendar);
    bool runDelta(shared_ptr<Calendar> calendar, const string & startURL, const GraphDeltaState & window, bool full);
    const json * seriesMaster(shared_ptr<Calendar> calendar, const string & masterId);
    size_t applyDeltaPage(shared_ptr<Calendar> calendar, const json & items, set<string> * seenEventIds);
    void removeEvents(vector<string> eventIds);
    void removeEventsNotIn(const string & calendarId, const set<string> & keepEventIds);
    void removeCalendar(shared_ptr<Calendar> calendar);

public:
    GraphCalendarWorker(shared_ptr<Account> account);

    void run() override;
};

#endif /* GraphCalendarWorker_hpp */
