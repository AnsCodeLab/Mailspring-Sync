//
//  GraphCalendarWorker.cpp
//  MailSync
//
//  Use of this file is subject to the terms and conditions defined
//  in 'LICENSE.md', which is part of the Mailspring-Sync package.
//

#include "GraphCalendarWorker.hpp"

#include <chrono>
#include <exception>
#include <set>
#include <thread>

#include "DAVWorker.hpp"
#include "Event.hpp"
#include "GraphEventConverter.hpp"
#include "MailStoreTransaction.hpp"
#include "MailUtils.hpp"
#include "NetworkRequestUtils.hpp"
#include "SyncException.hpp"
#include "XOAuth2TokenManager.hpp"
#include "icalendar.h"

static const char * GRAPH_CALENDARS_URL = "https://graph.microsoft.com/v1.0/me/calendars";
static const int GRAPH_MAX_ATTEMPTS = 4;
static const size_t GRAPH_DB_CHUNK = 100;

// Response headers of the request in flight on this thread, for Retry-After.
static thread_local string graphResponseHeaders;

static size_t graphHeaderCallback(char * buffer, size_t size, size_t nitems, void *) {
    graphResponseHeaders.append(buffer, size * nitems);
    return size * nitems;
}

static string graphRetryAfterHeader() {
    string lower = graphResponseHeaders;
    transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
    size_t pos = lower.find("retry-after:");
    if (pos == string::npos) return "";
    pos += strlen("retry-after:");
    size_t end = graphResponseHeaders.find_first_of("\r\n", pos);
    string value = graphResponseHeaders.substr(pos, end == string::npos ? string::npos : end - pos);
    size_t first = value.find_first_not_of(" \t");
    return first == string::npos ? "" : value.substr(first);
}

static string stringValue(const json & j, const char * key) {
    if (!j.is_object()) return "";
    auto it = j.find(key);
    return (it != j.end() && it->is_string()) ? it->get<string>() : "";
}

static bool hasStatus(const SyncException & e, const char * code) {
    return e.key.find(code) != string::npos;
}

// A 401/403 on the token request or on Graph itself means the credentials, not one
// resource, are bad: every other calendar would fail the same way.
static bool isAuthFailure(const SyncException & e) {
    return hasStatus(e, "Code: 401") || e.debuginfo.find("invalid_grant") != string::npos ||
           (hasStatus(e, "Code: 403") && e.debuginfo.find("ErrorAccessDenied") == string::npos);
}

// Per-calendar permission or existence problems (a shared calendar whose owner revoked
// access, a calendar deleted since the list was fetched) must not block the other calendars.
static bool isCalendarSpecificFailure(const SyncException & e) {
    return hasStatus(e, "Code: 404") ||
           (hasStatus(e, "Code: 403") && e.debuginfo.find("ErrorAccessDenied") != string::npos);
}

GraphCalendarWorker::GraphCalendarWorker(shared_ptr<Account> account) :
    store(new MailStore()),
    account(account),
    logger(spdlog::get("logger"))
{
}

void GraphCalendarWorker::run() {
    syncCalendarList();
}

// GET with a Graph access token, retrying 429/503 after a bounded Retry-After wait.
json GraphCalendarWorker::graphGet(const string & url) {
    for (int attempt = 1;; attempt++) {
        auto token = SharedXOAuth2TokenManager()->partsForAccount(account, OAuthResource::Graph);

        // Times in UTC so the converter never needs time zone rules.
        CURL * request = CreateJSONRequest(url, "GET", "Bearer " + token.accessToken, nullptr,
                                           {"Prefer: outlook.timezone=\"UTC\", odata.maxpagesize=100"});
        curl_easy_setopt(request, CURLOPT_TIMEOUT, 120L);
        curl_easy_setopt(request, CURLOPT_HEADERFUNCTION, graphHeaderCallback);
        graphResponseHeaders.clear();

        try {
            return PerformJSONRequest(request);
        } catch (SyncException & e) {
            bool throttled = hasStatus(e, "Code: 429") || hasStatus(e, "Code: 503");
            if (!throttled || attempt >= GRAPH_MAX_ATTEMPTS) {
                throw;
            }
            int waitSec = GraphBoundedRetryAfterSec(graphRetryAfterHeader());
            logger->warn("Graph request throttled ({}); retrying in {}s", e.key, waitSec);
            this_thread::sleep_for(chrono::seconds(waitSec));
        }
    }
}

void GraphCalendarWorker::syncCalendarList() {
    vector<json> remote;
    string url = GRAPH_CALENDARS_URL;
    while (!url.empty()) {
        json page = graphGet(url);
        if (!page.is_object() || !page.contains("value") || !page["value"].is_array()) {
            throw SyncException("graph-bad-response", "Calendar list had no value array: " + page.dump().substr(0, 500), true);
        }
        for (auto & item : page["value"]) {
            if (!stringValue(item, "id").empty()) remote.push_back(item);
        }
        url = stringValue(page, "@odata.nextLink");
    }

    auto local = store->findAllMap<Calendar>(Query().equal("accountId", account->id()), "id");
    set<string> remoteIds;
    vector<shared_ptr<Calendar>> toSync;

    for (auto & item : remote) {
        string graphId = stringValue(item, "id");
        string id = MailUtils::idForCalendar(account->id(), graphId);
        remoteIds.insert(id);

        string name = stringValue(item, "name");
        string color = GraphCalendarColor(stringValue(item, "hexColor"), stringValue(item, "color"));
        bool readOnly = !(item.contains("canEdit") && item["canEdit"].is_boolean() && item["canEdit"].get<bool>());

        shared_ptr<Calendar> calendar = local.count(id) ? local[id] : nullptr;
        bool changed = false;
        if (!calendar) {
            calendar = make_shared<Calendar>(id, account->id());
            calendar->setPath(graphId);
            changed = true;
        }
        if (changed || calendar->name() != name) {
            calendar->setName(name);
            changed = true;
        }
        string previousColor = calendar->color();
        calendar->setColor(color);
        if (changed || calendar->color() != previousColor) {
            changed = true;
        }
        if (changed || calendar->readOnly() != readOnly) {
            calendar->setReadOnly(readOnly);
            changed = true;
        }
        if (changed) {
            store->save(calendar.get());
        }
        toSync.push_back(calendar);
    }

    if (remote.empty()) {
        // Every mailbox has a default calendar; an empty list is a bad answer, not a request
        // to delete everything the user has stored.
        logger->warn("Graph returned no calendars; keeping the {} stored ones", local.size());
    } else {
        for (auto & pair : local) {
            if (!remoteIds.count(pair.first)) {
                logger->info("Calendar '{}' no longer exists on the server; removing it", pair.second->name());
                removeCalendar(pair.second);
            }
        }
    }

    exception_ptr pending = nullptr;
    for (auto & calendar : toSync) {
        try {
            syncCalendar(calendar);
        } catch (SyncException & e) {
            if (isAuthFailure(e)) throw;
            if (isCalendarSpecificFailure(e)) {
                logger->warn("Skipping calendar '{}': {}", calendar->name(), e.toJSON().dump());
            } else if (!pending) {
                pending = current_exception();
            }
        }
    }
    if (pending) rethrow_exception(pending);
}

void GraphCalendarWorker::syncCalendar(shared_ptr<Calendar> calendar) {
    auto range = getCalendarSyncRange();
    string token = calendar->syncToken();

    if (!GraphNeedsFullResync(token, range.start)) {
        GraphDeltaState state;
        GraphDeltaStateFromToken(token, state);
        try {
            runDelta(calendar, state.deltaLink, state, false);
            return;
        } catch (SyncException & e) {
            // syncStateNotFound: the delta token expired server-side.
            if (!hasStatus(e, "Code: 410")) throw;
            logger->info("Delta token for calendar '{}' expired; resyncing in full", calendar->name());
        }
    }

    GraphDeltaState window;
    window.windowStart = range.start;
    window.windowEnd = range.end;
    runDelta(calendar, GraphDeltaStartURL(calendar->path(), window.windowStart, window.windowEnd), window, true);
}

// Follows the delta pages from `startURL`. The new token is stored only after the last page is
// applied, so an interrupted pass is repeated rather than skipped.
void GraphCalendarWorker::runDelta(shared_ptr<Calendar> calendar, const string & startURL, const GraphDeltaState & window, bool full) {
    set<string> seen;
    string next = startURL;
    string deltaLink;
    size_t items = 0;

    while (!next.empty()) {
        json page = graphGet(next);
        if (!page.is_object() || !page.contains("value") || !page["value"].is_array()) {
            throw SyncException("graph-bad-response", "Delta page had no value array: " + page.dump().substr(0, 500), true);
        }
        items += page["value"].size();
        applyDeltaPage(calendar, page["value"], full ? &seen : nullptr);

        next = stringValue(page, "@odata.nextLink");
        if (next.empty()) {
            deltaLink = stringValue(page, "@odata.deltaLink");
        }
    }
    if (deltaLink.empty()) {
        throw SyncException("graph-bad-response", "Delta response ended without a deltaLink", true);
    }

    if (full) {
        removeEventsNotIn(calendar->id(), seen);
    }

    GraphDeltaState state = window;
    state.deltaLink = deltaLink;
    calendar->setSyncToken(GraphDeltaStateToToken(state));
    store->save(calendar.get());
    logger->info("Calendar '{}': {} sync applied {} change(s)", calendar->name(), full ? "full" : "delta", items);
}

namespace {
struct DeltaOp {
    bool remove = false;
    string graphId;
    string eventId;          // for removals
    string etag;
    string ics;              // CRLF text that is stored
    shared_ptr<ICalendar> parsed;
};
}

void GraphCalendarWorker::applyDeltaPage(shared_ptr<Calendar> calendar, const json & items, set<string> * seenEventIds) {
    // Convert and parse outside the transaction so the write lock is held only for the saves.
    vector<DeltaOp> ops;
    for (auto & item : items) {
        DeltaOp op;
        op.graphId = stringValue(item, "id");
        if (op.graphId.empty()) continue;

        if (item.contains("@removed")) {
            op.remove = true;
            op.eventId = MailUtils::idForEvent(account->id(), calendar->id(), GraphEventUID(op.graphId), "");
            ops.push_back(op);
            continue;
        }

        try {
            op.ics = GraphEventToICS(item);
        } catch (invalid_argument & e) {
            logger->warn("Skipping Graph event {}: {}", op.graphId, e.what());
            continue;
        }
        op.etag = stringValue(item, "changeKey");
        op.parsed = make_shared<ICalendar>(ICSStripCR(op.ics));
        if (op.parsed->Events.empty() || op.parsed->Events.front()->DtStart.IsEmpty()) {
            logger->warn("Skipping Graph event {}: generated ICS has no start", op.graphId);
            continue;
        }
        ops.push_back(op);
    }

    MailStoreTransaction transaction{store.get(), "GraphCalendarWorker:applyDeltaPage"};
    for (auto & op : ops) {
        if (op.remove) {
            auto existing = store->find<Event>(Query().equal("id", op.eventId).equal("calendarId", calendar->id()));
            if (existing) store->remove(existing.get());
            continue;
        }

        ICalendarEvent * icsEvent = op.parsed->Events.front();
        string eventId = MailUtils::idForEvent(account->id(), calendar->id(), icsEvent->UID, icsEvent->RecurrenceId);
        if (seenEventIds) seenEventIds->insert(eventId);

        auto existing = store->find<Event>(Query().equal("id", eventId));
        if (existing) {
            // Full resyncs re-deliver everything; don't re-persist (and re-notify) what is unchanged.
            if (existing->etag() == op.etag && existing->icsData() == op.ics) continue;
            existing->applyICSEventData(op.etag, op.graphId, op.ics, icsEvent);
            store->save(existing.get());
        } else {
            Event event(op.etag, account->id(), calendar->id(), op.ics, icsEvent);
            event.setHref(op.graphId);
            store->save(&event);
        }
    }
    transaction.commit();
}

void GraphCalendarWorker::removeEvents(vector<string> eventIds) {
    for (auto chunk : MailUtils::chunksOfVector(eventIds, GRAPH_DB_CHUNK)) {
        MailStoreTransaction transaction{store.get(), "GraphCalendarWorker:removeEvents"};
        // Per-model removal so Event::afterRemove clears the EventSearch row too.
        for (auto & event : store->findAll<Event>(Query().equal("id", chunk))) {
            store->remove(event.get());
        }
        transaction.commit();
    }
}

void GraphCalendarWorker::removeEventsNotIn(const string & calendarId, const set<string> & keepEventIds) {
    vector<string> stale;
    SQLite::Statement select(store->db(), "SELECT id FROM Event WHERE calendarId = ?");
    select.bind(1, calendarId);
    while (select.executeStep()) {
        string id = select.getColumn(0).getString();
        if (!keepEventIds.count(id)) stale.push_back(id);
    }
    if (!stale.empty()) {
        logger->info("Removing {} event(s) no longer in the sync window", stale.size());
        removeEvents(stale);
    }
}

void GraphCalendarWorker::removeCalendar(shared_ptr<Calendar> calendar) {
    removeEventsNotIn(calendar->id(), {});
    MailStoreTransaction transaction{store.get(), "GraphCalendarWorker:removeCalendar"};
    store->remove(calendar.get());
    transaction.commit();
}
