---
id: S04
parent: M001
milestone: M001
provides:
  - npm-publish-ready package (files whitelist, engines, keywords, license, repository, .npmignore, executable shebang)
  - MIT LICENSE file
  - English README documenting zero-config workflow, all 25 tools, Claude Desktop config, troubleshooting
  - Actionable error messages for AADSTS auth codes and HTTP 401/403/404 responses
  - 16 new tests covering error message paths (72 total, up from 56)
requires:
  - slice: S03
    provides: All 25 tools validated with new auth layer, dual-audience token routing
affects: []
key_files:
  - package.json
  - LICENSE
  - .npmignore
  - README.md
  - src/auth.js
  - src/client.js
  - tests/auth.test.js
  - tests/client.test.js
key_decisions:
  - files whitelist uses src/ directory (not individual files) to automatically include future source files
  - Injectable fetchFn in SharePointClient constructor for testable HTTP error paths (extends existing pattern from discoverTenantId)
  - Exported wrapAuthError as pure function from auth.js for independent unit testing without MSAL mocks
patterns_established:
  - npm pack --dry-run as canonical publish verification
  - AADSTS code detection via regex on both error.errorCode and error.message properties
  - Tool reference table in README organized by category mirrors src/tools.js grouping
observability_surfaces:
  - "npm pack --dry-run — lists exactly what would be published (8 files, 12.1kB)"
  - "node src/index.js — emits startup line on stderr confirming bin+shebang+permissions"
  - "Error messages include actionable guidance — AADSTS codes and HTTP status codes produce distinct remediation text"
  - "Original error details preserved in parentheses for developer debugging"
drill_down_paths:
  - .gsd/milestones/M001/slices/S04/tasks/T01-SUMMARY.md
  - .gsd/milestones/M001/slices/S04/tasks/T02-SUMMARY.md
  - .gsd/milestones/M001/slices/S04/tasks/T03-SUMMARY.md
duration: 33m
verification_result: passed
completed_at: 2026-03-16
---

# S04: npm Package & Polish

**Package is publish-ready with npm metadata, MIT license, English README documenting zero-config workflow for all 25 tools, and actionable error messages for auth and API failures.**

## What Happened

Three tasks assembled the final packaging layer on top of the validated auth+tools stack from S01–S03.

**T01 (npm publish-readiness):** Added `files`, `engines`, `keywords`, `license`, `repository` to package.json. Created MIT LICENSE. Created `.npmignore` as belt-and-suspenders exclusion. Set executable permission on `src/index.js`. Verified with `npm pack --dry-run` (8 files, no leaks) and `node src/index.js` (startup message on stderr).

**T02 (README rewrite):** Replaced the German README (which documented the old Azure app registration + env var workflow) with a complete English rewrite. Covers: one-liner quick start with `npx sharepoint-online-mcp`, Claude Desktop JSON config with zero env vars, how the auth works (well-known client ID, device code flow, auto tenant discovery, token caching), tool reference table with all 25 tool names organized by category, requirements, known limitations (Conditional Access), and troubleshooting table. Verified no German fragments, no env var references, no Azure app registration steps, all 25 tool names present.

**T03 (error guidance):** Added `wrapAuthError()` pure function to `src/auth.js` mapping MSAL errors to user-facing guidance — specific AADSTS codes (50076, 53003, 700016, 50059) get targeted messages, others get generic message with code preserved. Enhanced both `request()` and `spRest()` in `src/client.js` with status-specific messages for 401 (auth expired), 403 (access denied), 404 (not found). Made fetch injectable via `options.fetchFn` for testability. Added 16 new tests (8 auth, 8 client).

## Verification

- **npm pack --dry-run:** 8 files — LICENSE, README.md, package.json, src/auth-cli.js, src/auth.js, src/client.js, src/index.js, src/tools.js. No tests, .gsd, .git, .env. ✅
- **Server startup:** `node src/index.js` emits `🚀 SharePoint Online MCP Server started (stdio)`. ✅
- **Test suite:** 72 tests pass, 0 fail (56 original + 16 new error message tests). ✅
- **Tool count invariant:** `grep -c 'server.tool(' src/tools.js` = 25. ✅
- **README validation:** All 25 tool names present, zero German fragments, zero env var references, zero Azure app registration references. ✅

## Requirements Advanced

- R013 — Package is now npm-publish-ready: correct `files`, `engines`, `bin`, `keywords`, `license`, `repository`. `npm pack --dry-run` confirms clean artifact. `npx .` starts server.
- R015 — Auth errors (AADSTS50076, 53003, 700016, 50059) and client errors (401, 403, 404) now produce actionable guidance messages with remediation steps.
- R014 — Confirmed no secrets in published package. `npm pack --dry-run` shows only source code and docs.

## Requirements Validated

- R013 — `npm pack --dry-run` produces clean 8-file artifact. `node src/index.js` starts successfully. Package metadata complete.
- R015 — 16 tests prove specific AADSTS codes and HTTP status codes produce actionable error messages. `wrapAuthError` and status-specific client errors verified at unit level.

## New Requirements Surfaced

- none

## Requirements Invalidated or Re-scoped

- none

## Deviations

- T03 added injectable `fetchFn` to SharePointClient constructor — not in original plan but necessary for testing HTTP error paths without module-level mocking. Follows the existing injectable pattern from `discoverTenantId`. No breaking changes.
- Test count grew from 56 to 72 (plan expected 56 baseline preserved — 72 exceeds that).

## Known Limitations

- Conditional Access blocking documented in README as known limitation — no runtime fallback (R016 deferred).
- Package is publish-ready but not yet published to npm — actual `npm publish` is a manual step.
- All tool validation is contract-level (mock tests). Live SharePoint tenant testing is milestone-level UAT, not slice-level.

## Follow-ups

- Milestone-level UAT: test full workflow against a live SharePoint tenant (start server → authenticate → search sites → create page → publish).
- Consider `npm publish` once milestone UAT passes.

## Files Created/Modified

- `package.json` — Added license, engines, files, keywords, repository fields
- `LICENSE` — New MIT license file (2026, Daniel Laurin)
- `.npmignore` — New exclusion list for npm publish
- `src/index.js` — chmod +x (content unchanged)
- `README.md` — Complete rewrite from German to English zero-config documentation
- `src/auth.js` — Added `wrapAuthError()` export function; wrapped device code flow catch block
- `src/client.js` — Added injectable `fetchFn` in constructor; status-specific error messages in `request()` and `spRest()`
- `tests/auth.test.js` — Added 8 `wrapAuthError` tests
- `tests/client.test.js` — Added 8 HTTP error message tests (4 Graph, 4 SP REST)

## Forward Intelligence

### What the next slice should know
- This is the terminal slice of M001. The milestone is code-complete pending live UAT. All 25 tools are validated at contract level, package is publish-ready, docs are current.

### What's fragile
- The well-known client ID (`d3590ed6-...`) scope availability has not been tested against a live tenant with Graph Beta pages endpoints. If Microsoft restricts scopes on this client ID, the auth layer works but specific tool calls may get 403s.
- README tool table is manually maintained — any tool additions/removals in src/tools.js require a corresponding README update.

### Authoritative diagnostics
- `npm pack --dry-run` — definitive view of what gets published, catches any leaked files
- `node --test tests/*.test.js` — 72 tests, covers auth error wrapping, client error paths, all 25 tool handlers
- `grep -c 'server.tool(' src/tools.js` — tool count invariant, must stay at 25

### What assumptions changed
- Original plan assumed 56 tests as baseline — actual count is 72 after error message test additions. The extra 16 tests cover error paths that previously had no coverage.
