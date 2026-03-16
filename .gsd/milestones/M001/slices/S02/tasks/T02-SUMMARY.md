---
id: T02
parent: S02
milestone: M001
provides:
  - connect_to_site MCP tool with automatic tenant discovery and site resolution from any SharePoint URL
  - parseSharePointUrl exported utility handling sites/teams/root/subpage/trailing-slash edge cases
  - tests/tools.test.js with 7 URL parsing test cases
key_files:
  - src/tools.js
  - tests/tools.test.js
key_decisions:
  - Pass empty string (not "/") to getSiteByUrl for root sites — Graph endpoint sites/{host}:/ resolves correctly with empty sitePath
  - First cross-module import in tools.js: discoverTenantId from auth.js
patterns_established:
  - Exported pure functions for testable parsing logic, tool handler wires parsing + async calls
observability_surfaces:
  - connect_to_site error paths include original URL and upstream error messages (tenant discovery HTTP status, Graph API status)
duration: 10m
verification_result: passed
completed_at: 2026-03-16
blocker_discovered: false
---

# T02: Add `connect_to_site` tool with URL parsing and tests

**Added `connect_to_site` MCP tool and `parseSharePointUrl` export with 7 passing URL parsing tests**

## What Happened

Added `parseSharePointUrl(url)` as an exported function at the top of `src/tools.js`. It extracts `hostname` and `sitePath` from any SharePoint URL, handling `/sites/`, `/teams/`, root site (empty path), trailing slashes, subpage paths, and ports. Invalid URLs throw with the original URL in the message.

Registered `connect_to_site` tool between `get_site_by_url` and `list_my_sites`. It chains: parse URL → `discoverTenantId(hostname)` → `getSiteByUrl(hostname, sitePath)` → return site details + tenant ID. Error path returns `isError: true` with the upstream message.

Added `import { discoverTenantId } from "./auth.js"` — first cross-module import in tools.js beyond zod.

Created `tests/tools.test.js` with 7 test cases covering standard site, root site, trailing slash, subpage path, teams prefix, invalid URL, and URL with port.

## Verification

- `node --test tests/tools.test.js` → 7/7 pass
- `node --test tests/auth.test.js` → 12/12 pass
- `node --test tests/client.test.js` → 6/6 pass
- `grep -c 'server.tool(' src/tools.js` → 25
- `node src/index.js` → starts clean ("SharePoint Online MCP Server started")
- `grep -E '[äöüÄÖÜß]' src/tools.js` → no matches (no German text)

### Slice-level verification status

| Check | Status |
|-------|--------|
| `node --test tests/tools.test.js` all pass | ✅ |
| URL parsing: valid site, root, trailing slash, subpage, teams, invalid | ✅ |
| `node --test tests/auth.test.js && tests/client.test.js` 18 existing pass | ✅ |
| `node src/index.js` starts without error | ✅ |
| `grep -c 'server.tool(' src/tools.js` = 25 | ✅ |
| No German text in tools.js | ✅ |
| Failure path: `connect_to_site` returns actionable error for invalid URL | ✅ (parseSharePointUrl throws "Invalid SharePoint URL: {url}") |
| Tool registration: `connect_to_site` and `list_my_sites` registered | ✅ |

All slice-level verification checks pass.

## Diagnostics

- `connect_to_site` errors propagate upstream messages: invalid URL → "Invalid SharePoint URL: {url}", tenant discovery failure → domain + HTTP status, Graph API failure → "API {status}: {body}"
- Tool list visible via MCP `tools/list` — 25 tools total
- `parseSharePointUrl` is exported and directly testable

## Deviations

- Plan suggested `sitePath || "/"` for root site handling. Changed to pass `sitePath` as empty string directly — `getSiteByUrl(hostname, "")` produces `sites/{hostname}:/` which is the correct Graph API path for root sites. Using `"/"` would produce `sites/{hostname}://` (double slash, incorrect).

## Known Issues

None

## Files Created/Modified

- `src/tools.js` — Added `parseSharePointUrl` export, `discoverTenantId` import from `./auth.js`, `connect_to_site` tool registration (25 total tools)
- `tests/tools.test.js` — New: 7 test cases for URL parsing edge cases
