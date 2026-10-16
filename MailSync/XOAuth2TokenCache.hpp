//
//  XOAuth2TokenCache.hpp
//  MailSync
//
//  Use of this file is subject to the terms and conditions defined
//  in 'LICENSE.md', which is part of the Mailspring-Sync package.
//

#ifndef XOAuth2TokenCache_hpp
#define XOAuth2TokenCache_hpp

#include <ctime>
#include <map>
#include <string>

// An OAuth refresh token mints separate access tokens per resource audience: Microsoft
// accepts the mail (IMAP/SMTP) token only at outlook.office.com and the Graph token only at
// graph.microsoft.com, so the two must be cached side by side.
enum class OAuthResource {
    Mail,
    Graph,
};

struct XOAuth2Parts {
    std::string username;
    std::string accessToken;
    time_t expiryDate;
};

inline std::string XOAuth2CacheKey(const std::string & accountId, OAuthResource resource) {
    return accountId + (resource == OAuthResource::Graph ? "|graph" : "|mail");
}

// Not thread safe: XOAuth2TokenManager serialises access with its own lock.
class XOAuth2TokenCache {
    std::map<std::string, XOAuth2Parts> _entries;

public:
    // Returns true and fills `out` if the entry has more than `bufferSec` seconds left at `now`.
    bool lookup(const std::string & accountId, OAuthResource resource, time_t now, int bufferSec, XOAuth2Parts & out) const {
        auto it = _entries.find(XOAuth2CacheKey(accountId, resource));
        if (it == _entries.end() || it->second.expiryDate <= now + bufferSec) {
            return false;
        }
        out = it->second;
        return true;
    }

    void store(const std::string & accountId, OAuthResource resource, const XOAuth2Parts & parts) {
        _entries[XOAuth2CacheKey(accountId, resource)] = parts;
    }
};

#endif /* XOAuth2TokenCache_hpp */
