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
#include <mutex>
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

// Log/diagnostic dumps include HTTP bodies, which need not be valid UTF-8.
static string safeDump(const json & j) {
    return j.dump(-1, ' ', false, json::error_handler_t::replace);
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

// The store is deliberately never freed, like the other workers': the long-lived instance is
// used by a thread that is still running when the process exits, and a destructor racing that
// thread would close the database under it.
GraphCalendarWorker::GraphCalendarWorker(shared_ptr<Account> account) :
    store(new MailStore()),
    account(account),
    logger(spdlog::get("logger"))
{
}

void GraphCalendarWorker::run() {
    // The periodic loop and the sync-calendar trigger each own a worker; two concurrent passes
    // would race on the same calendars.
    static mutex passLock;
    lock_guard<mutex> guard(passLock);
    syncCalendarList();
}

// GET with a Graph access token, retrying 429/503 after a bounded Retry-After wait.
json GraphCalendarWorker::graphGet(const string & url) {
    if (!GraphIsTrustedURL(url)) {
        // The bearer token must never leave graph.microsoft.com.
        throw SyncException("graph-untrusted-url", "Refusing to send Graph credentials to " + url.substr(0, 200), false);
    }

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
            throw SyncException("graph-bad-response", "Calendar list had no value array: " + safeDump(page).substr(0, 500), true);
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
        // Event writes are rejected for these accounts (see TaskProcessor), so every calendar
        // is stored read-only whatever Graph's canEdit says; the UI then offers no edits that
        // would fail.
        if (changed || !calendar->readOnly()) {
            calendar->setReadOnly(true);
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
    bool warnedAccessDenied = false;
    for (auto & calendar : toSync) {
        try {
            syncCalendar(calendar);
        } catch (SyncException & e) {
            if (isAuthFailure(e)) throw;
            if (isCalendarSpecificFailure(e)) {
                if (hasStatus(e, "Code: 403") && !warnedAccessDenied) {
                    warnedAccessDenied = true;
                    logger->warn("Microsoft Graph denied access (ErrorAccessDenied) to calendar '{}'; skipping it. "
                                 "If this affects calendars that should be readable, the Calendars.ReadWrite.Shared "
                                 "permission may not be granted for this account.", calendar->name());
                } else {
                    logger->info("Skipping calendar '{}': {}", calendar->name(), safeDump(e.toJSON()));
                }
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

    GraphDeltaState state;
    bool needsFull = true;
    if (!GraphNeedsFullResync(token, range.start) && GraphDeltaStateFromToken(token, state) && GraphIsTrustedURL(state.deltaLink)) {
        try {
            needsFull = runDelta(calendar, state.deltaLink, state, false);
        } catch (SyncException & e) {
            // 410 syncStateNotFound: the delta token expired server-side; 400: Graph rejects it.
            if (!GraphStatusMeansStaleDeltaToken(e.key)) throw;
            logger->info("Delta token for calendar '{}' rejected ({}); resyncing in full", calendar->name(), e.key);
        }
    }
    if (!needsFull) return;

    GraphDeltaState window;
    window.windowStart = range.start;
    window.windowEnd = range.end;
    runDelta(calendar, GraphDeltaStartURL(calendar->path(), window.windowStart, window.windowEnd), window, true);
}

// Follows the delta pages from `startURL`. The new token is stored only after the last page is
// applied, so an interrupted pass is repeated rather than skipped. Returns true when the pass
// saw a series master and, being incremental, left it to the caller to resync in full.
bool GraphCalendarWorker::runDelta(shared_ptr<Calendar> calendar, const string & startURL, const GraphDeltaState & window, bool full) {
    set<string> seen;
    masterCache.clear();
    sawSeriesMaster = false;
    passMasterItems = 0;
    string next = startURL;
    string deltaLink;
    size_t items = 0;
    size_t skipped = 0;

    while (!next.empty()) {
        json page = graphGet(next);
        if (!page.is_object() || !page.contains("value") || !page["value"].is_array()) {
            throw SyncException("graph-bad-response", "Delta page had no value array: " + safeDump(page).substr(0, 500), true);
        }
        items += page["value"].size();
        skipped += applyDeltaPage(calendar, page["value"], full ? &seen : nullptr);

        next = stringValue(page, "@odata.nextLink");
        if (next.empty()) {
            deltaLink = stringValue(page, "@odata.deltaLink");
        }
    }
    if (deltaLink.empty()) {
        throw SyncException("graph-bad-response", "Delta response ended without a deltaLink", true);
    }
    if (!GraphIsTrustedURL(deltaLink)) {
        throw SyncException("graph-untrusted-url", "Delta response returned a deltaLink outside graph.microsoft.com", false);
    }
    if (GraphPassUnusable(items - passMasterItems, skipped)) {
        // Neither purge nor advance the token: the next pass retries from the same point.
        throw SyncException("graph-bad-response", "Calendar '" + calendar->name() + "': all " + to_string(items - passMasterItems) + " received event(s) were unusable", true);
    }
    if (GraphShouldResyncAfterPass(full, sawSeriesMaster)) {
        // An edited series: its occurrences hold a merged copy of the old master. The caller
        // runs a full resync, which saves the token; nothing is saved here.
        logger->info("Calendar '{}': a recurring series changed; resyncing in full", calendar->name());
        return true;
    }

    if (full) {
        removeEventsNotIn(calendar->id(), seen);
    }

    GraphDeltaState state = window;
    state.deltaLink = deltaLink;
    calendar->setSyncToken(GraphDeltaStateToToken(state));
    store->save(calendar.get());
    logger->info("Calendar '{}': {} sync applied {} change(s)", calendar->name(), full ? "full" : "delta", items);
    return false;
}

// Master of a series, from this pass's delta items or fetched once per pass; nullptr when the
// server will not give it out (404/403), in which case the occurrence is stored as delivered.
// Any other failure propagates and fails the pass.
const json * GraphCalendarWorker::seriesMaster(shared_ptr<Calendar> calendar, const string & masterId) {
    auto cached = masterCache.find(masterId);
    if (cached != masterCache.end()) {
        return cached->second.is_object() ? &cached->second : nullptr;
    }
    json fetched;
    try {
        fetched = graphGet(GraphEventURL(calendar->path(), masterId));
    } catch (SyncException & e) {
        if (!hasStatus(e, "Code: 404") && !hasStatus(e, "Code: 403")) throw;
        logger->warn("Series master {} is unavailable ({}); storing its occurrences without it", masterId, e.key);
        masterCache[masterId] = json(nullptr);
        return nullptr;
    }
    if (!fetched.is_object() || stringValue(fetched, "id").empty()) {
        masterCache[masterId] = json(nullptr);
        return nullptr;
    }
    return &(masterCache[masterId] = fetched);
}

namespace {
struct DeltaOp {
    bool remove = false;
    string graphId;
    string eventId;          // for removals
    string etag;
    string ics;              // CRLF text that is stored
    shared_ptr<ICalendar> parsed;
    bool hasUtcBounds = false;
    long long rs = 0;
    long long re = 0;
};

// applyICSEventData derives rs/re with Date::toUnix(), which reads a `Z` time as local time.
// For timed events the true UTC instants replace them; all-day events keep the local-midnight
// bounds the app's day windows expect.
void applyUtcBounds(Event & event, const DeltaOp & op) {
    if (!op.hasUtcBounds) return;
    event._data["rs"] = op.rs;
    event._data["re"] = op.re;
}
}

// Returns how many non-removal items could not be converted and were skipped.
size_t GraphCalendarWorker::applyDeltaPage(shared_ptr<Calendar> calendar, const json & items, set<string> * seenEventIds) {
    // Convert and parse outside the transaction so the write lock is held only for the saves.
    vector<DeltaOp> ops;
    size_t skipped = 0;
    // Masters first so occurrences earlier on the page can use them without a request.
    for (auto & item : items) {
        if (item.contains("@removed") || !GraphIsSeriesMaster(item)) continue;
        string masterId = stringValue(item, "id");
        if (masterId.empty()) continue;
        masterCache[masterId] = item;
        sawSeriesMaster = true;
        passMasterItems++;
    }
    for (auto & item : items) {
        DeltaOp op;
        op.graphId = stringValue(item, "id");
        if (op.graphId.empty()) continue;
        // The master is a data source only (cached above); storing it would show the first
        // occurrence twice.
        if (!item.contains("@removed") && GraphIsSeriesMaster(item)) continue;

        string eventId = MailUtils::idForEvent(account->id(), calendar->id(), GraphEventUID(op.graphId), "");
        if (item.contains("@removed")) {
            op.remove = true;
            op.eventId = eventId;
            ops.push_back(op);
            continue;
        }

        // Recorded before conversion: an event the server still lists but this build cannot
        // convert must survive a full resync's purge of events that are gone.
        if (seenEventIds) seenEventIds->insert(eventId);

        // Occurrences/exceptions arrive without their series' descriptive fields (see
        // GraphMergeSeriesMaster); a master that cannot be read leaves them as delivered.
        json merged;
        const json * source = &item;
        if (GraphNeedsSeriesMerge(item)) {
            const json * master = seriesMaster(calendar, stringValue(item, "seriesMasterId"));
            if (master != nullptr) {
                merged = GraphMergeSeriesMaster(item, *master);
                source = &merged;
            }
        }

        vector<string> warnings;
        try {
            op.ics = GraphEventToICS(*source, &warnings);
        } catch (invalid_argument & e) {
            logger->warn("Skipping Graph event {}: {}", op.graphId, e.what());
            skipped++;
            continue;
        }
        for (auto & warning : warnings) {
            logger->warn("Graph event {}: {}", op.graphId, warning);
        }
        op.etag = stringValue(item, "changeKey");
        op.parsed = make_shared<ICalendar>(ICSStripCR(op.ics));
        if (op.parsed->Events.empty() || op.parsed->Events.front()->DtStart.IsEmpty()) {
            logger->warn("Skipping Graph event {}: generated ICS has no start", op.graphId);
            skipped++;
            continue;
        }
        op.hasUtcBounds = GraphEventUnixBounds(*source, op.rs, op.re);
        ops.push_back(op);
    }

    MailStoreTransaction transaction{store, "GraphCalendarWorker:applyDeltaPage"};
    for (auto & op : ops) {
        if (op.remove) {
            auto existing = store->find<Event>(Query().equal("id", op.eventId).equal("calendarId", calendar->id()));
            if (existing) store->remove(existing.get());
            continue;
        }

        ICalendarEvent * icsEvent = op.parsed->Events.front();
        string eventId = MailUtils::idForEvent(account->id(), calendar->id(), icsEvent->UID, icsEvent->RecurrenceId);

        auto existing = store->find<Event>(Query().equal("id", eventId));
        if (existing) {
            // Full resyncs re-deliver everything; don't re-persist (and re-notify) what is unchanged.
            if (existing->etag() == op.etag && existing->icsData() == op.ics) continue;
            existing->applyICSEventData(op.etag, op.graphId, op.ics, icsEvent);
            applyUtcBounds(*existing, op);
            store->save(existing.get());
        } else {
            Event event(op.etag, account->id(), calendar->id(), op.ics, icsEvent);
            event.setHref(op.graphId);
            applyUtcBounds(event, op);
            store->save(&event);
        }
    }
    transaction.commit();
    return skipped;
}

void GraphCalendarWorker::removeEvents(vector<string> eventIds) {
    for (auto chunk : MailUtils::chunksOfVector(eventIds, GRAPH_DB_CHUNK)) {
        MailStoreTransaction transaction{store, "GraphCalendarWorker:removeEvents"};
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
    MailStoreTransaction transaction{store, "GraphCalendarWorker:removeCalendar"};
    store->remove(calendar.get());
    transaction.commit();
}
