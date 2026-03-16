# S02: Site Discovery & Connection Tools

**Goal:** Claude can search SharePoint sites, resolve a site from any URL with automatic tenant discovery, list followed sites, and disconnect/reconnect — all via MCP tools with English descriptions.
**Demo:** User gives Claude a SharePoint URL → `connect_to_site` parses it, discovers tenant, resolves site ID → `search_sites` finds other sites → `list_my_sites` shows followed sites → all tool descriptions are in English.

## Must-Haves

- `connect_to_site` MCP tool: takes a full SharePoint URL, parses hostname/path, calls `discoverTenantId`, resolves site via `getSiteByUrl`, returns site details
- `list_my_sites` MCP tool: returns the user's followed sites via `/me/followedSites`
- All 22 existing German tool descriptions translated to English
- URL parsing handles edge cases: root sites, trailing slashes, subpage paths beyond site path, `/teams/` prefix
- Unit tests for URL parsing logic and new tool registrations

## Proof Level

- This slice proves: integration (tools wire auth + client + tenant discovery into working MCP endpoints)
- Real runtime required: no (unit tests with mocks prove wiring; live SharePoint deferred to UAT)
- Human/UAT required: no (deferred to milestone UAT)

## Verification

- `node --test tests/tools.test.js` — new test file, all tests pass:
  - URL parsing: valid site URL, root site, trailing slash, subpage path, teams prefix, invalid URL
  - Tool registration: `connect_to_site` and `list_my_sites` are registered with correct schemas
- `node --test tests/auth.test.js && node --test tests/client.test.js` — existing 18 tests still pass (no regressions)
- `node src/index.js` starts without error
- `grep -c 'server.tool(' src/tools.js` returns 25 (23 existing + 2 new)
- `grep -P '[äöüÄÖÜß]' src/tools.js` returns empty (no German text remains)
- At least one verification check for a failure path: `connect_to_site` returns actionable error for invalid URL input

## Observability / Diagnostics

- Runtime signals: `connect_to_site` logs tenant discovery result to stderr before site resolution; errors include the original URL for context
- Inspection surfaces: tool list visible via MCP `tools/list` protocol message
- Failure visibility: invalid URL → error with "Invalid SharePoint URL" + the URL; tenant discovery failure → error with domain + HTTP status (inherited from `discoverTenantId`)
- Redaction constraints: none (SharePoint URLs are not secrets)

## Integration Closure

- Upstream surfaces consumed: `discoverTenantId` from `src/auth.js`, `getSiteByUrl` / `listSites` from `src/client.js`, `registerTools` pattern from `src/tools.js`
- New wiring introduced: `connect_to_site` is the first tool in `tools.js` that imports from `auth.js` (cross-module import for `discoverTenantId`)
- What remains before milestone is truly usable end-to-end: S03 (validate all 20+ tools with new auth), S04 (npm packaging, README, error polish)

## Tasks

- [x] **T01: Translate tool descriptions to English and add `list_my_sites`** `est:25m`
  - Why: German descriptions make the MCP server unusable for non-German speakers. `list_my_sites` is a trivial wrapper that pairs naturally with this cleanup pass — both are mechanical changes to `tools.js`.
  - Files: `src/tools.js`
  - Do: Translate all 22 German tool descriptions to clear English. Add `list_my_sites` tool that calls `client.listSites("")` (which hits `/me/followedSites`) with no parameters, returns array of `{id, name, url, description}`. Keep tool registration order logical: site tools grouped together.
  - Verify: `grep -P '[äöüÄÖÜß]' src/tools.js` returns empty; `node src/index.js` starts without error; `grep -c 'server.tool(' src/tools.js` returns 24 (23 + 1 new)
  - Done when: zero German text in tools.js, `list_my_sites` tool registered, server starts clean

- [ ] **T02: Add `connect_to_site` tool with URL parsing and tests** `est:35m`
  - Why: This is the primary new capability — a single-URL entry point that combines URL parsing → tenant discovery → auth → site resolution. It also needs the slice's test file covering URL parsing edge cases and tool registration.
  - Files: `src/tools.js`, `tests/tools.test.js` (new)
  - Do: (1) Add `connect_to_site` tool to `tools.js` that takes a single `url` string parameter. Implementation: parse with `new URL(url)`, extract hostname, extract site path (first two path segments after `/` — handles `sites/name`, `teams/name`, or empty for root), call `discoverTenantId(hostname)` (import from `../auth.js`), call `client.getSiteByUrl(hostname, sitePath)`, return site details. Handle edge cases: strip trailing slashes, ignore path segments beyond site path, handle root site (empty path), return actionable error for invalid URLs. (2) Create `tests/tools.test.js` with unit tests: extract the URL parsing logic into a testable helper (either inline in the test or exported from tools.js — prefer a small exported `parseSharePointUrl(url)` function since URL parsing is the risk). Test cases: standard site URL, root site, trailing slash, subpage path, teams prefix, invalid URL, non-SharePoint URL. Add tool registration smoke test using a mock server pattern or simple import check.
  - Verify: `node --test tests/tools.test.js` all pass; `node --test tests/auth.test.js && node --test tests/client.test.js` still pass; `grep -c 'server.tool(' src/tools.js` returns 25; `node src/index.js` starts without error
  - Done when: `connect_to_site` tool registered with correct schema, URL parsing handles all edge cases with tests proving it, all existing tests still pass

## Files Likely Touched

- `src/tools.js` — English translations, `list_my_sites` tool, `connect_to_site` tool, `parseSharePointUrl` export
- `tests/tools.test.js` — new test file for URL parsing and tool registration
