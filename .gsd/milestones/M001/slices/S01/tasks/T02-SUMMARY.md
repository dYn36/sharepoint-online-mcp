---
id: T02
parent: S01
milestone: M001
provides:
  - dual-audience token routing in SharePointClient (Graph vs SP REST)
  - zero-config server startup (no env var gates)
  - disconnect MCP tool for cache clearing and re-auth
  - zero-config auth-cli standalone tester
  - package.json with all runtime dependencies declared
key_files:
  - src/client.js
  - src/index.js
  - src/tools.js
  - src/auth-cli.js
  - tests/client.test.js
  - package.json
key_decisions:
  - Used new URL(siteUrl).origin for SP REST hostname extraction — handles all subdomain variants (contoso, contoso-admin, contoso-my) without string manipulation
patterns_established:
  - Test pattern for client token routing: mock auth records getAccessToken args, let fetch fail naturally since token acquisition happens before fetch — verifies routing without needing fetch mocks
observability_surfaces:
  - Server startup message on stderr confirms zero-config mode
  - Disconnect tool provides explicit user-facing feedback on cache clear
  - Auth errors include resource URL that failed
  - SP REST errors include the SharePoint hostname attempted
duration: 15m
verification_result: passed
completed_at: 2026-03-16
blocker_discovered: false
---

# T02: Wire dual-audience tokens through client and enable zero-config server startup

**Connected zero-config auth engine to client, index, and tools — server starts with no env vars, routes per-resource tokens, and exposes disconnect tool.**

## What Happened

Updated 5 source files and created 1 test file to complete the S01 zero-config auth wiring:

1. **client.js** — `request()` now passes `"https://graph.microsoft.com"` to `getAccessToken()`, `spRest()` extracts origin via `new URL(siteUrl).origin` and passes it as the resource. All 20+ existing API methods unchanged.

2. **index.js** — Removed all `process.env.SHAREPOINT_*` reads and the error block. Server constructs `new SharePointAuth()` with no args, passes `auth` to `registerTools`. Description and startup message switched to English.

3. **tools.js** — `registerTools` signature extended to `(server, client, auth)`. Added `disconnect` tool that calls `auth.logout()` and returns confirmation text. All existing tools untouched.

4. **auth-cli.js** — Stripped env var dependencies, constructs zero-config auth, calls `getAccessToken("https://graph.microsoft.com")`. Messages in English.

5. **package.json** — Renamed to `sharepoint-online-mcp`, bin key updated, dependencies block added with `@azure/msal-node ^5.1.0`, `@modelcontextprotocol/sdk ^1.6.1`, `node-fetch ^3.3.2`, `zod ^3.24.1`.

6. **tests/client.test.js** — 6 tests covering Graph resource routing via `request()` and `graphBetaReq()`, SP REST origin extraction for standard/admin/my subdomains, and path stripping verification.

## Verification

- `node --test tests/client.test.js` — **6/6 pass** (dual-audience routing)
- `node --test tests/auth.test.js` — **12/12 pass** (T01 tests still green)
- Server startup: `node src/index.js` emits `🚀 SharePoint Online MCP Server started (stdio)` on stderr, no crash, no env var errors
- `grep -rn "process\.env\.SHAREPOINT" src/ | grep -v "\/\/"` — **zero hits** (no env var references)
- `node -e "import('./src/tools.js').then(m => console.log(typeof m.registerTools))"` — prints `function`

### Slice-level verification status (S01):
- ✅ `node --test tests/auth.test.js` — 12/12 pass
- ✅ `node --test tests/client.test.js` — 6/6 pass
- ✅ Server starts without crash, no env var error on stderr
- ✅ `grep -rn "process\.env\.SHAREPOINT" src/` — zero hits
- ⏳ Manual: device code flow → token → Graph API call (requires human verification)

## Diagnostics

- **Server startup**: `node src/index.js` — stderr shows startup message immediately, then device code prompt on first tool call
- **Token cache**: `~/.sharepoint-mcp-cache.json` existence confirms prior auth
- **Auth state on stderr**: silent token acquisition failures logged with resource URL, device code prompt shows verification URI
- **Disconnect**: The `disconnect` MCP tool clears cache and resets auth state; next request triggers fresh device code flow

## Deviations

- Ran `npm install` in worktree to install newly declared dependencies (`@modelcontextprotocol/sdk`, `node-fetch`, `zod`) — they weren't in node_modules since the original codebase only had `@azure/msal-node` installed.
- `@azure/msal-node` version kept at `^5.1.0` (matching installed) instead of plan's `^2.16.2` which is an older major version.

## Known Issues

- None

## Files Created/Modified

- `src/client.js` — dual-audience token routing in `request()` and `spRest()`
- `src/index.js` — zero-config startup, English messages, passes `auth` to `registerTools`
- `src/tools.js` — `disconnect` tool added, `registerTools` signature extended with `auth` param
- `src/auth-cli.js` — zero-config standalone auth tester
- `package.json` — renamed to `sharepoint-online-mcp`, runtime dependencies declared
- `tests/client.test.js` — 6 unit tests for dual-audience token routing
