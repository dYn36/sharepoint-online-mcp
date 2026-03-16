---
id: T01
parent: S01
milestone: M001
provides:
  - zero-config SharePointAuth class (no constructor args)
  - discoverTenantId(domain) exported function for tenant GUID resolution
  - buildScopes(resource) exported helper for per-resource scope construction
  - test infrastructure (node --test, tests/ directory)
key_files:
  - src/auth.js
  - tests/auth.test.js
  - package.json
key_decisions:
  - Used Node built-in fetch (v25) instead of node-fetch for discoverTenantId — fewer deps
  - Extracted buildScopes as exported pure function for testability and reuse by client.js
  - Made discoverTenantId accept optional fetchFn parameter for test injection
  - Added stderr logging on silent token acquisition failure for observability
patterns_established:
  - Injectable fetch pattern for testable network calls (fetchFn parameter with default)
  - node:test + node:assert/strict as project test framework
  - Resource-scoped token acquisition via buildScopes helper
observability_surfaces:
  - stderr: "[auth] Silent token acquisition failed for {resource}: {message}" on silent acquire failure
  - stderr: device code prompt (English) with verification URI and user code
  - All discoverTenantId errors include domain, HTTP status, and failure reason
  - ~/.sharepoint-mcp-cache.json existence indicates prior successful auth
duration: ~15min
verification_result: passed
completed_at: 2026-03-16
blocker_discovered: false
---

# T01: Refactored auth.js to zero-config multi-resource auth engine with 12 unit tests

**Rewrote auth.js from parameterized (clientId/tenantId) to zero-config using Microsoft Office well-known client ID, added tenant discovery and per-resource scope construction, established test framework.**

## What Happened

Rewrote `src/auth.js` from a 95-line parameterized auth class to a zero-config multi-resource auth engine:

1. **Removed constructor parameters** — `SharePointAuth()` takes zero args. Uses hardcoded well-known client ID `d3590ed6-52b3-4102-aeff-aad2292ab01c` and `common` authority.

2. **Added `discoverTenantId(domain, fetchFn?)`** — Calls OpenID config endpoint, parses tenant GUID from `token_endpoint` URL. Uses Node built-in `fetch` (not `node-fetch`). Accepts injectable `fetchFn` for testing. Error messages include domain, HTTP status, and specific failure reason.

3. **Added `buildScopes(resource)`** — Pure helper that converts resource URL to `["{resource}/.default"]` scopes array. Handles trailing slashes. Extracted as standalone export for testability and reuse by `client.js` in T02.

4. **Changed `getAccessToken(scopes)` → `getAccessToken(resource)`** — Now accepts a resource URL, delegates to `buildScopes`. Removed `DEFAULT_SCOPES` constant entirely.

5. **Translated device code prompt to English** — Was German ("SharePoint Login erforderlich!"), now English.

6. **Added observability** — Silent acquisition failures now emit to stderr with resource URL and error message.

7. **Established test framework** — Added `"test": "node --test"` to package.json. Created `tests/auth.test.js` with 12 tests across 3 suites using `node:test` + `node:assert/strict`.

8. **Installed `@azure/msal-node`** — Required for import resolution in tests. Was listed as implicit dep but never declared/installed.

## Verification

- `node --test tests/auth.test.js` — **12/12 pass** (3 buildScopes, 6 discoverTenantId, 3 SharePointAuth construction)
- `node -e "import('./src/auth.js').then(m => { console.log(typeof m.SharePointAuth, typeof m.discoverTenantId, typeof m.buildScopes) })"` — prints `function function function`
- `grep "d3590ed6" src/auth.js` — well-known client ID present
- `grep "DEFAULT_SCOPES" src/auth.js` — zero hits (removed)
- `grep "constructor()" src/auth.js` — zero-arg constructor confirmed
- `grep "login required" src/auth.js` — English device code message confirmed
- `grep "process.env" src/auth.js` — zero hits (no env vars in auth module)

### Slice-level verification status (partial — T01 is first of 2 tasks):
- ✅ `node --test tests/auth.test.js` — passes
- ⏳ `node --test tests/client.test.js` — not yet created (T02)
- ⏳ Server start without crash — depends on T02 (index.js still reads env vars)
- ⏳ `grep -rn "process.env.SHAREPOINT" src/` — 4 hits remain in index.js and auth-cli.js (T02 scope)
- ⏳ Manual device code flow — human verification deferred

## Diagnostics

- **Test suite**: `node --test tests/auth.test.js` — 12 tests, all pure/mock-based, no network calls
- **Auth state on stderr**: silent failures log `[auth] Silent token acquisition failed for {resource}`, device code prompt shows verification URI
- **Token cache**: `~/.sharepoint-mcp-cache.json` — existence means prior auth succeeded
- **Error shapes**: All `discoverTenantId` errors are `Error` with message containing domain attempted + specific failure reason (network error, HTTP status, missing field, unparseable GUID)

## Deviations

- **Used Node built-in `fetch` instead of `node-fetch`** — Plan noted auth.js imports `node-fetch`, but the original code didn't actually use fetch at all. Since Node 25 has stable global `fetch`, used that directly. Simpler, fewer deps.
- **Installed `@azure/msal-node` as npm dependency** — Wasn't in package.json `dependencies` (no dependencies block existed). Required for tests to import auth.js. T02 will add the full dependencies block.
- **Added `buildScopes` as a third export** — Plan suggested extracting scope construction to a helper. Made it a named export for reuse by client.js.

## Known Issues

- None

## Files Created/Modified

- `src/auth.js` — Complete rewrite: zero-config constructor, well-known client ID, discoverTenantId, buildScopes, resource-scoped getAccessToken, English device code prompt
- `tests/auth.test.js` — New: 12 unit tests across 3 suites (buildScopes, discoverTenantId, SharePointAuth)
- `package.json` — Added `"test": "node --test"` script; `@azure/msal-node` added to dependencies by npm install
