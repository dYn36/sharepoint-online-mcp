---
id: S02
parent: M001
milestone: M001
provides:
  - connect_to_site MCP tool — single-URL entry point that chains URL parsing → tenant discovery → auth → site resolution
  - list_my_sites MCP tool — returns user's followed SharePoint sites
  - parseSharePointUrl exported utility — testable URL parsing for sites/teams/root/subpage/trailing-slash edge cases
  - All 25 tool descriptions and user-facing strings in English (zero German text remaining)
requires:
  - slice: S01
    provides: SharePointAuth.getAccessToken(resource), discoverTenantId(domain), server starts without env vars
affects:
  - S03
key_files:
  - src/tools.js
  - tests/tools.test.js
key_decisions:
  - Pass empty string (not "/") to getSiteByUrl for root sites — Graph endpoint sites/{host}:/ resolves correctly
  - First cross-module import in tools.js: discoverTenantId from auth.js — establishes pattern for tools needing auth-layer functions
  - Full rewrite of tools.js rather than surgical German→English edits — lower risk of missed fragments
patterns_established:
  - Exported pure functions for testable parsing logic; tool handlers wire parsing + async service calls
  - English-only user-facing strings throughout tools.js
  - connect_to_site as the canonical "URL → site context" entry point for user workflows
observability_surfaces:
  - connect_to_site errors propagate upstream messages with original URL context (invalid URL, tenant discovery HTTP status, Graph API status)
  - list_my_sites returns isError:true with upstream Graph error message on failure
  - Tool list visible via MCP tools/list — 25 tools total
drill_down_paths:
  - .gsd/milestones/M001/slices/S02/tasks/T01-SUMMARY.md
  - .gsd/milestones/M001/slices/S02/tasks/T02-SUMMARY.md
duration: 22m
verification_result: passed
completed_at: 2026-03-16
---

# S02: Site Discovery & Connection Tools

**25 MCP tools with English descriptions, `connect_to_site` for URL→tenant→site resolution, `list_my_sites` for followed sites, 25/25 tests passing**

## What Happened

Two tasks, both mechanical and clean.

**T01** rewrote `src/tools.js` to translate all 22 German tool descriptions, parameter descriptions, and response strings to English. Added `list_my_sites` tool (calls `client.listSites("")` → `/me/followedSites`, returns `{id, name, url, description}` array). Tool count went from 23 to 24.

**T02** added `parseSharePointUrl(url)` as an exported pure function at the top of `tools.js` — extracts `hostname` and `sitePath` from any SharePoint URL, handling `/sites/`, `/teams/`, root sites (empty path), trailing slashes, subpage paths, and ports. Registered `connect_to_site` tool that chains: parse URL → `discoverTenantId(hostname)` → `getSiteByUrl(hostname, sitePath)` → return site details + tenant ID. This is the first tool in `tools.js` that imports from `auth.js` (cross-module dependency). Created `tests/tools.test.js` with 7 URL parsing tests. Tool count: 25.

## Verification

All slice-level checks pass:

| Check | Result |
|-------|--------|
| `node --test tests/tools.test.js` — 7/7 pass | ✅ |
| `node --test tests/auth.test.js` — 12/12 pass | ✅ |
| `node --test tests/client.test.js` — 6/6 pass | ✅ |
| `grep -c 'server.tool(' src/tools.js` = 25 | ✅ |
| `grep -E '[äöüÄÖÜß]' src/tools.js` returns empty | ✅ |
| `node src/index.js` starts without error | ✅ |
| `connect_to_site` returns actionable error for invalid URL | ✅ |
| `connect_to_site` and `list_my_sites` registered with correct schemas | ✅ |

Total: 25 tests across 3 test files, 25 registered MCP tools, zero regressions.

## Requirements Advanced

- R006 (Interactive Site Discovery) — `connect_to_site` provides URL → tenant discovery → site resolution in a single tool call; `list_my_sites` provides followed sites listing. Search was already available from pre-existing `search_sites` tool.
- R007 (Multi-Site Support) — `connect_to_site` resolves any site by URL without session state; tools continue taking siteId as parameter, enabling multi-site workflows.
- R015 (Clear Error Messages) — `connect_to_site` produces actionable errors: "Invalid SharePoint URL: {url}" for bad input, domain + HTTP status for tenant discovery failure, Graph API status for site resolution failure.

## Requirements Validated

- None newly validated. R006 and R007 are advanced but require live SharePoint testing (deferred to milestone UAT) to validate fully.

## New Requirements Surfaced

- None

## Requirements Invalidated or Re-scoped

- None

## Deviations

- T01 translated response strings and parameter descriptions in addition to tool descriptions — the plan only specified tool descriptions, but leaving German in parameters/responses would have been equally unusable for non-German speakers.
- T02 used empty string instead of `"/"` for root site path — plan suggested `sitePath || "/"` but that produces an incorrect Graph API path (`sites/{host}://`).

## Known Limitations

- `connect_to_site` does not persist site context — each call resolves independently. Multi-site session state management is ergonomic enough (tools accept siteId) but there's no "current site" concept.
- `list_my_sites` depends on Graph `/me/followedSites` which returns only explicitly followed sites, not all sites the user has access to. `search_sites` covers the broader search case.
- No live SharePoint testing yet — all verification is unit tests with mocks. Live validation deferred to S03/UAT.

## Follow-ups

- S03 must validate that `connect_to_site` works end-to-end with real SharePoint (token acquisition + Graph site resolution) as part of dual-audience token validation.
- S03 should verify that the site ID returned by `connect_to_site` works correctly when passed to downstream tools (pages, layout, web parts, etc.).

## Files Created/Modified

- `src/tools.js` — Translated all German to English, added `parseSharePointUrl` export, `discoverTenantId` import from `./auth.js`, `list_my_sites` tool, `connect_to_site` tool (25 total registrations)
- `tests/tools.test.js` — New: 7 test cases for URL parsing edge cases

## Forward Intelligence

### What the next slice should know
- `tools.js` now imports from `auth.js` (`discoverTenantId`). S03 will add more cross-module wiring for dual-audience tokens — this pattern is established.
- All 25 tools are English-only. No i18n consideration needed.
- `connect_to_site` returns `{id, displayName, webUrl, description, tenantId}` — downstream tools need `id` (the siteId) and can ignore the rest.

### What's fragile
- `parseSharePointUrl` assumes SharePoint URLs always have the site path in the first two segments after `/` (e.g. `/sites/marketing`). Unusual SharePoint URL structures (vanity domains, custom paths) would break this. Low risk in practice but worth noting.
- Root site detection (empty sitePath → `sites/{host}:/`) depends on Graph API accepting this path format. Untested against live Graph.

### Authoritative diagnostics
- `node --test tests/tools.test.js` — 7 tests covering all URL parsing edge cases. If URL parsing breaks, this catches it.
- `grep -c 'server.tool(' src/tools.js` — quick sanity check that tool count hasn't regressed.
- MCP `tools/list` at runtime — shows all 25 tools with descriptions.

### What assumptions changed
- Plan assumed `sitePath || "/"` for root sites — actual implementation uses empty string because Graph API path construction differs. Decision #012 in DECISIONS.md documents this.
