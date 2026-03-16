# S02: Site Discovery & Connection Tools — Research

**Date:** 2026-03-16

## Summary

S02 adds site discovery and connection tools to the MCP server so Claude can search for SharePoint sites, resolve a site from a URL (with automatic tenant discovery), list sites the user follows, and manage connections. The existing codebase already has most of the building blocks — `search_sites`, `get_site_by_url`, and `get_site_details` tools exist, `discoverTenantId` is exported from auth.js, and `disconnect` was added in S01. The main new work is a `connect_to_site` tool that combines URL parsing → tenant discovery → auth → site resolution into a single user-friendly flow, a `list_my_sites` tool wrapping the followed-sites endpoint, and translating existing German tool descriptions to English.

This is straightforward wiring of established patterns. No new libraries, no risky integration, no ambiguous scope.

## Recommendation

Add `connect_to_site` as the primary new tool — it takes a full SharePoint URL (e.g. `https://contoso.sharepoint.com/sites/marketing`), parses hostname and path, calls `discoverTenantId(hostname)`, triggers auth lazily via `client.getSiteByUrl()`, and returns site details. Add `list_my_sites` as a simple wrapper around `GET /me/followedSites`. Translate all existing German tool descriptions to English for consistency. Add a client method for `/me/followedSites` that doesn't require a search term (the existing `listSites("")` falls through to followed sites, but a dedicated method is cleaner). Add unit tests for the URL parsing logic in `connect_to_site` and for the new/updated client methods.

## Implementation Landscape

### Key Files

- `src/tools.js` (710 lines) — All MCP tool registrations. Has existing `search_sites`, `get_site_by_url`, `get_site_details`. Needs: new `connect_to_site` tool, new `list_my_sites` tool, English descriptions on all existing tools.
- `src/client.js` (320 lines) — API client. Has `listSites(search)` which falls through to `/me/followedSites` when search is empty, `getSiteByUrl(hostname, sitePath)`, `getSite(siteId)`. Needs: potentially a dedicated `getFollowedSites()` method for clarity, though `listSites("")` already works.
- `src/auth.js` (175 lines) — Exports `discoverTenantId(domain)` and `buildScopes(resource)`. `connect_to_site` tool will import and use `discoverTenantId` directly. No changes needed.
- `src/index.js` (40 lines) — Server entrypoint, already passes `auth` to `registerTools`. No changes needed.
- `tests/client.test.js` — Existing 6 tests. Add tests for new client methods or verify existing ones cover the needed paths.

### Tool Inventory (current → target)

| Tool | Status | Changes needed |
|------|--------|----------------|
| `search_sites` | Exists, German description | Translate to English |
| `get_site_details` | Exists, German description | Translate to English |
| `get_site_by_url` | Exists, German description | Translate to English |
| `connect_to_site` | **New** | URL parsing → tenant discovery → auth → site resolution |
| `list_my_sites` | **New** | Wrapper around `/me/followedSites` |
| `disconnect` | Exists, English | No changes (already done in S01) |
| All other tools | Exist, German descriptions | Translate to English |

### `connect_to_site` Design

Takes a single `url` parameter (e.g. `https://contoso.sharepoint.com/sites/marketing`). Internally:

1. `new URL(url)` to parse — extract `hostname` and `pathname`
2. Split pathname: first two segments after `/` form the site path (e.g. `sites/marketing`)
3. Call `discoverTenantId(hostname)` — triggers HTTP call, returns tenant GUID
4. Call `client.getSiteByUrl(hostname, sitePath)` — triggers Graph API auth lazily
5. Return site ID, name, URL, description — same shape as `get_site_by_url` but from a single URL input

Edge cases to handle:
- Root site URL (no `/sites/` path) — site path would be empty or `/`
- URL with trailing slash
- URL with subpages beyond the site path (e.g. `/sites/marketing/SitePages/Home.aspx`)
- Invalid URL format

The `discoverTenantId` call is informational for the user flow — it proves the tenant is reachable. The actual auth happens lazily inside `client.getSiteByUrl()` when it calls `this.auth.getAccessToken()`.

### `list_my_sites` Design

No parameters. Calls `client.listSites("")` which hits `GET /me/followedSites`. Returns array of `{id, name, url, description}`. Simple.

### Build Order

1. **Translate all German tool descriptions to English** — mechanical, no risk, cleans up the codebase for consistency. Do this first to avoid merge conflicts with new tool additions.
2. **Add `connect_to_site` tool** — the primary new capability. Imports `discoverTenantId` from auth.js. Adds URL parsing logic.
3. **Add `list_my_sites` tool** — trivial wrapper.
4. **Add tests** — URL parsing edge cases for `connect_to_site`, tool registration verification.

### Verification Approach

- `node --test` — all existing tests (18) still pass after changes
- New tests: URL parsing logic for `connect_to_site` (valid URLs, root site, trailing slashes, subpages, invalid input)
- `node src/index.js` starts without error, emits startup message on stderr
- `grep` to verify no German descriptions remain in tools.js
- Tool count verification: confirm all expected tools are registered

## Constraints

- `tools.js` is already 710 lines — adding tools pushes it further. Stay under 800 if possible. The file is organized by clear section headers which keeps it navigable.
- `connect_to_site` must import `discoverTenantId` from `../auth.js` — tools.js currently only imports `zod`. This is the first cross-module import in tools.js.
- URL parsing for SharePoint site paths: the standard format is `/{type}/{name}` where type is `sites`, `teams`, or just `/` for root. The path segments beyond the site are not part of the site identifier.

## Common Pitfalls

- **SharePoint URL path parsing** — `/sites/marketing/SitePages/Home.aspx` needs to extract just `sites/marketing`, not the full path. The Graph API `sites/{hostname}:/{sitePath}` endpoint expects exactly the site-relative path. Over-including path segments returns 404.
- **Root site has no path** — `https://contoso.sharepoint.com` is a valid site URL with an empty site path. The Graph endpoint for root site is `sites/{hostname}` (no colon path syntax).
