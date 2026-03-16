---
id: S01
parent: M001
milestone: M001
provides:
  - zero-config SharePointAuth class (well-known client ID, no constructor args, no env vars)
  - discoverTenantId(domain) for automatic tenant GUID resolution from SharePoint URLs
  - buildScopes(resource) helper for per-resource scope construction
  - dual-audience token routing in SharePointClient (Graph API vs SP REST API)
  - disconnect MCP tool for cache clearing and re-authentication
  - zero-config server startup (no env var gates, no config files)
  - package.json with correct name, bin, and all runtime dependencies
  - test infrastructure (node:test, 18 unit tests across 2 suites)
requires:
  - slice: none
    provides: first slice — no upstream dependencies
affects:
  - S02
  - S03
  - S04
key_files:
  - src/auth.js
  - src/client.js
  - src/index.js
  - src/tools.js
  - src/auth-cli.js
  - tests/auth.test.js
  - tests/client.test.js
  - package.json
key_decisions:
  - Well-known Microsoft Office client ID d3590ed6-52b3-4102-aeff-aad2292ab01c hardcoded — no env vars, no app registration
  - Node built-in fetch for discoverTenantId instead of node-fetch — fewer deps, Node 25 has stable global fetch
  - Injectable fetchFn parameter pattern for testable network calls without global mocking
  - new URL(siteUrl).origin for SP REST hostname extraction — handles all subdomain variants cleanly
  - node:test + node:assert/strict as project test framework — zero test deps
  - buildScopes extracted as named export for reuse across auth.js and client.js
patterns_established:
  - Zero-config class pattern: hardcode well-known constants, derive everything else at runtime
  - Injectable fetch pattern: optional fetchFn parameter with default to global fetch for test injection
  - Resource-scoped token acquisition: getAccessToken(resource) builds scopes internally via buildScopes
  - Mock-auth test pattern for client: mock auth records getAccessToken args, let fetch fail naturally to verify routing without fetch mocks
  - English-only user-facing messages on stderr (replaced German originals)
observability_surfaces:
  - stderr "[auth] Silent token acquisition failed for {resource}: {message}" on silent acquire failure
  - stderr device code prompt with verification URI and user code (transient, not logged)
  - stderr "🚀 SharePoint Online MCP Server started (stdio)" on successful startup
  - ~/.sharepoint-mcp-cache.json existence indicates prior successful auth
  - All discoverTenantId errors include domain, HTTP status, and failure reason
  - disconnect tool provides explicit user-facing confirmation on cache clear
drill_down_paths:
  - .gsd/milestones/M001/slices/S01/tasks/T01-SUMMARY.md
  - .gsd/milestones/M001/slices/S01/tasks/T02-SUMMARY.md
duration: ~30min
verification_result: passed
completed_at: 2026-03-16
---

# S01: Zero-Config Auth Engine

**Server starts with zero env vars, authenticates via device code using well-known client ID, routes per-resource tokens for Graph and SP REST APIs, and caches tokens across restarts.**

## What Happened

Two tasks rebuilt the auth layer and wired it through the entire codebase:

**T01** rewrote `src/auth.js` from a parameterized class (required clientId/tenantId) to a zero-config engine. The `SharePointAuth` class now takes zero constructor arguments — it hardcodes Microsoft Office's well-known client ID `d3590ed6-52b3-4102-aeff-aad2292ab01c` and uses `common` authority. Added `discoverTenantId(domain)` that resolves tenant GUIDs from SharePoint domains via the OpenID configuration endpoint. Added `buildScopes(resource)` that converts resource URLs to `["{resource}/.default"]` scope arrays. Established the test framework with `node:test` and wrote 12 unit tests covering scope construction, tenant discovery (with injected fetch mocks), and auth class construction. Device code prompt translated from German to English.

**T02** connected the new auth engine to the rest of the codebase. `client.js` now routes tokens per-resource: `request()` passes `"https://graph.microsoft.com"` for Graph calls, `spRest()` extracts the SharePoint origin via `new URL(siteUrl).origin` for SP REST calls. `index.js` was stripped of all `process.env.SHAREPOINT_*` reads and error blocks — the server now starts immediately with zero configuration. A `disconnect` MCP tool was added to `tools.js` that clears the token cache and enables re-authentication. `auth-cli.js` was updated for zero-config standalone testing. `package.json` was renamed to `sharepoint-online-mcp` with all runtime dependencies declared. 6 client tests verify dual-audience routing across Graph, standard, admin, and OneDrive SharePoint subdomains.

## Verification

All slice-level automated checks pass:

- `node --test tests/auth.test.js` — **12/12 pass** (buildScopes ×3, discoverTenantId ×6, SharePointAuth ×3)
- `node --test tests/client.test.js` — **6/6 pass** (Graph routing ×2, SP REST origin extraction ×3, path stripping ×1)
- Server startup: `node src/index.js` emits startup message on stderr, no crash, no env var error
- `grep -rn "process\.env\.SHAREPOINT" src/ | grep -v "\/\/"` — **zero hits**

Manual verification deferred: device code flow → token → Graph API call requires human interaction with a live Azure AD tenant.

## Requirements Advanced

- R001 (Zero-config Auth via Well-Known Client ID) — auth.js hardcodes Office client ID, constructs with zero args, unit tested. Live tenant proof deferred to manual UAT.
- R002 (Auto Tenant Discovery) — discoverTenantId implemented and unit tested with 6 test cases. Live proof deferred to manual UAT.
- R003 (Device Code Flow) — implemented in auth.js with English prompt on stderr. Live proof deferred to manual UAT.
- R004 (Token Caching with Silent Renewal) — MSAL cache plugin writes to ~/.sharepoint-mcp-cache.json. Cross-restart persistence proof deferred to manual UAT.
- R005 (Logout/Disconnect Tool) — disconnect tool registered, calls auth.logout(). Live proof deferred to manual UAT.
- R015 (Clear Error Messages) — auth errors include resource URL, HTTP status, MSAL error code. Tenant discovery errors include domain and failure reason. Full error message review in S04.

## Requirements Validated

- R014 (No Hardcoded Secrets) — grep confirms zero `process.env.SHAREPOINT` references in src/. Well-known client ID is a public Microsoft constant, not a secret. No tenant IDs or credentials in source.

## New Requirements Surfaced

- none

## Requirements Invalidated or Re-scoped

- none

## Deviations

- Used Node built-in `fetch` instead of `node-fetch` for `discoverTenantId` — original codebase imported node-fetch but Node 25 has stable global fetch. Simpler, one fewer dependency.
- `@azure/msal-node` pinned at `^5.1.0` (installed version) instead of plan's `^2.16.2` which referenced an older major version.
- `buildScopes` extracted as a third named export from auth.js — not in original plan but enables clean reuse by client.js.

## Known Limitations

- **No live Azure AD proof yet** — all auth verification is unit-level with mocks. Real token acquisition, device code flow, and API calls require manual testing against a live tenant (covered in UAT).
- **Dual-audience token acquisition untested end-to-end** — client routes correct resources to getAccessToken, but whether MSAL actually returns valid tokens for both Graph and SP REST audiences with the well-known client ID is unproven until live testing.
- **Conditional Access blocking** — tenants that block device code flow will fail with no fallback (R016, deferred).
- **No site discovery tools yet** — S02 scope. Auth triggers lazily but there's no `connect_to_site` tool to trigger it.

## Follow-ups

- S02 needs to build site discovery tools that exercise the auth engine against real SharePoint
- S03 needs to verify all 20+ existing tools work with well-known client ID tokens (scope availability is the key risk)
- If well-known client ID lacks required Graph Beta scopes for pages API, S03 will need to try alternative client IDs (Azure CLI, VS Code)

## Files Created/Modified

- `src/auth.js` — Complete rewrite: zero-config constructor, well-known client ID, discoverTenantId, buildScopes, resource-scoped getAccessToken, English device code prompt, stderr observability
- `src/client.js` — Dual-audience token routing: request() → Graph resource, spRest() → SP origin extraction
- `src/index.js` — Zero-config startup: removed all env var reads, English messages, passes auth to registerTools
- `src/tools.js` — disconnect tool added, registerTools signature extended with auth param
- `src/auth-cli.js` — Zero-config standalone auth tester
- `tests/auth.test.js` — 12 unit tests: buildScopes, discoverTenantId (mocked fetch), SharePointAuth construction
- `tests/client.test.js` — 6 unit tests: Graph routing, SP REST origin extraction for standard/admin/my subdomains
- `package.json` — Renamed to sharepoint-online-mcp, test script, bin entry, all runtime dependencies declared

## Forward Intelligence

### What the next slice should know
- `SharePointAuth` is instantiated once in `index.js` and passed to both `SharePointClient` and `registerTools`. To get a token, call `auth.getAccessToken("https://graph.microsoft.com")` or `auth.getAccessToken("https://contoso.sharepoint.com")`. Auth triggers device code flow lazily on first call.
- `discoverTenantId(domain)` is a standalone exported function — S02's `connect_to_site` tool should call it directly to resolve tenant from a SharePoint URL before triggering auth.
- `buildScopes(resource)` is exported and reusable — constructs `["{resource}/.default"]` from a resource URL.
- The `disconnect` tool already exists. S02 just needs to add `connect_to_site`, `search_sites`, and `list_my_sites`.

### What's fragile
- **Well-known client ID scope availability is unproven** — the Office client ID `d3590ed6` may not have pre-consented scopes for Graph Beta pages/layout endpoints. S03 is where this risk gets retired or forces a client ID change. If it fails, `auth.js` line ~10 is the only change point.
- **MSAL cache format** — `~/.sharepoint-mcp-cache.json` is written by MSAL's cache plugin. If MSAL major version changes the format, cache breaks silently (user just re-auths, not catastrophic).

### Authoritative diagnostics
- `node --test tests/auth.test.js && node --test tests/client.test.js` — 18 tests, all pure/mock-based, sub-4s total. If these break, auth wiring is broken.
- `grep -rn "process\.env\.SHAREPOINT" src/ | grep -v "\/\/"` — must return empty. Any hits mean env var gates leaked back in.
- `node src/index.js` stderr — should show startup message immediately, no errors. Device code prompt appears only on first tool call.

### What assumptions changed
- **node-fetch dependency** — assumed needed, but Node 25 global fetch is sufficient. node-fetch is still in package.json dependencies (some existing code may import it) but auth.js doesn't use it.
- **MSAL version** — plan referenced ^2.16.2, actual installed is ^5.1.0. Major version jump, API is compatible for our usage.
