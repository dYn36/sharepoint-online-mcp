---
id: T03
parent: S04
milestone: M001
provides:
  - Actionable error messages for AADSTS auth codes in token acquisition
  - Status-specific error guidance for 401/403/404 HTTP responses in Graph and SP REST paths
  - 16 new tests covering all error message variants
key_files:
  - src/auth.js
  - src/client.js
  - tests/auth.test.js
  - tests/client.test.js
key_decisions:
  - Made fetch injectable via constructor options in SharePointClient (options.fetchFn) to enable testing HTTP error paths without module-level mocking
  - Exported wrapAuthError as a pure function from auth.js so it can be unit tested directly without mocking MSAL
patterns_established:
  - AADSTS code detection via regex on both error.errorCode and error.message properties
  - Injectable fetchFn pattern in SharePointClient constructor (matches existing fetchFn pattern in discoverTenantId)
observability_surfaces:
  - Error messages now include actionable guidance — AADSTS codes produce specific remediation text, HTTP status codes produce distinct messages
  - Original error details preserved in parentheses for developer debugging
duration: 15m
verification_result: passed
completed_at: 2026-03-16
blocker_discovered: false
---

# T03: Wrap auth and client errors with actionable guidance

**Added actionable error messages for AADSTS auth codes (Conditional Access, app not found, tenant missing) and HTTP 401/403/404 responses in both Graph and SP REST API paths.**

## What Happened

Added `wrapAuthError()` to `src/auth.js` — a pure function that maps MSAL errors to user-facing guidance. Specific AADSTS codes (50076, 53003, 700016, 50059) get targeted messages; other AADSTS codes get a generic message with the code preserved; non-AADSTS errors get a generic auth failure message. All include a "disconnect" remediation hint. The device code flow catch block now wraps errors through this function.

Enhanced both `request()` (Graph API) and `spRest()` (SP REST API) in `src/client.js` with status-specific error messages: 401 → authentication expired, 403 → access denied, 404 → resource not found. Other statuses get a generic "SharePoint API error" with status and body. Made `fetch` injectable via `options.fetchFn` constructor parameter to enable clean testing of error paths.

Added 8 tests in auth.test.js (wrapAuthError variants) and 8 tests in client.test.js (4 for Graph, 4 for SP REST error paths).

## Verification

- `node --test tests/*.test.js` — 72 tests pass (56 existing + 16 new), zero failures
- `grep -c 'server.tool(' src/tools.js` — 25 (tool count invariant preserved)
- `npm pack --dry-run` — only intended files included (8 files, 12.1kB)

### Slice-level verification status (T03 is final task):
- ✅ `npm pack --dry-run` — only src/, README.md, LICENSE, package.json included
- ✅ `node --test tests/*.test.js` — 72 tests pass (exceeds baseline of 56)
- ✅ `grep -c 'server.tool(' src/tools.js` — 25

## Diagnostics

- Grep for AADSTS codes in test output: `node --test tests/auth.test.js 2>&1 | grep -i aadsts`
- Grep for HTTP error messages in test output: `node --test tests/client.test.js 2>&1 | grep -i "error messages"`
- `wrapAuthError` is exported and can be called directly for diagnostic purposes

## Deviations

- Made `fetch` injectable via `SharePointClient` constructor `options.fetchFn` parameter — not in the original plan but necessary for testability. The existing client tests that catch-and-ignore fetch errors still work unchanged. This follows the same injectable pattern already used by `discoverTenantId(domain, fetchFn)`.

## Known Issues

None.

## Files Created/Modified

- `src/auth.js` — Added `wrapAuthError()` export function; wrapped device code flow catch block
- `src/client.js` — Added injectable `fetchFn` in constructor; replaced generic error throws with status-specific actionable messages in both `request()` and `spRest()`
- `tests/auth.test.js` — Added 8 `wrapAuthError` tests covering all AADSTS code paths and edge cases
- `tests/client.test.js` — Added 8 error message tests (4 Graph, 4 SP REST) using injectable `fetchFn`
