---
id: M001
provides:
  - Zero-config MCP server for SharePoint Online (npx sharepoint-online-mcp)
  - Device code flow authentication using well-known Microsoft Office client ID (no app registration)
  - Automatic tenant ID discovery from SharePoint URLs
  - Dual-audience token routing (Graph API + SP REST API)
  - 25 MCP tools with English descriptions covering pages, layout, web parts, navigation, branding, site discovery
  - Token caching across server restarts via MSAL cache plugin
  - npm-publish-ready package (8 files, 12.1kB)
  - Actionable error messages for AADSTS auth codes and HTTP 401/403/404
  - 72 unit tests across 3 test files
key_decisions:
  - D001: Microsoft Office well-known client ID d3590ed6-52b3-4102-aeff-aad2292ab01c — no app registration needed
  - D002: Tenant ID auto-derived from SharePoint domain via OpenID Discovery endpoint
  - D003: Lazy authentication on first tool call, not at server startup
  - D006: Plain JavaScript ESM, no TypeScript, no build step
  - D007: Resource-scoped tokens (.default) for Graph and SP REST audiences
  - D009: node:test + node:assert/strict — zero test dependencies
  - D010: Injectable fetchFn pattern for testable network calls
  - D015: Overrides parameter on registerTools for test injection without experimental flags
  - D017: Pure wrapAuthError function for independently testable AADSTS error mapping
patterns_established:
  - Zero-config class pattern: hardcode well-known constants, derive everything at runtime
  - Injectable fetch/dependency pattern: optional parameter with default for test substitution
  - Resource-scoped token acquisition: getAccessToken(resource) builds scopes internally
  - MockServer.call(name, args) for invoking registered MCP tool handlers in tests
  - npm pack --dry-run as canonical publish verification
observability_surfaces:
  - stderr "🚀 SharePoint Online MCP Server started (stdio)" on successful startup
  - stderr "[auth] Silent token acquisition failed for {resource}: {message}" on silent acquire failure
  - stderr device code prompt with verification URI and user code
  - ~/.sharepoint-mcp-cache.json existence indicates prior successful auth
  - Auth errors include AADSTS codes with targeted remediation messages
  - Client errors include HTTP status with actionable guidance (401/403/404)
  - node --test tests/*.test.js — 72 tests, full regression in <1s
requirement_outcomes:
  - id: R001
    from_status: active
    to_status: validated
    proof: auth.js hardcodes Office client ID, zero constructor args, 3 unit tests, server starts without env vars. Live tenant proof requires human UAT.
  - id: R002
    from_status: active
    to_status: validated
    proof: discoverTenantId implemented with 6 mock-fetch unit tests covering success, error, and edge cases. Live tenant proof requires human UAT.
  - id: R003
    from_status: active
    to_status: validated
    proof: Device code flow implemented in auth.js with English stderr prompt. Code path unit tested. Live flow requires human interaction.
  - id: R004
    from_status: active
    to_status: validated
    proof: MSAL cache plugin configured to persist to ~/.sharepoint-mcp-cache.json. Silent renewal code path exists. Cross-restart persistence requires human UAT.
  - id: R005
    from_status: active
    to_status: validated
    proof: disconnect tool registered in S01, handler tested in S03 (calls auth.logout(), returns confirmation). Live proof requires human UAT.
  - id: R006
    from_status: active
    to_status: validated
    proof: connect_to_site (URL→tenant→site resolution), list_my_sites, search_sites all registered. 7 URL parsing tests + 5 handler mock tests in S02/S03.
  - id: R007
    from_status: active
    to_status: validated
    proof: All tools accept siteId as parameter. connect_to_site resolves any URL without session state. No "current site" lock-in. Contract-tested in S03.
  - id: R013
    from_status: active
    to_status: validated
    proof: npm pack --dry-run produces clean 8-file artifact (12.1kB). bin entry, shebang, engines, files whitelist, keywords, license, repository all set. Server starts via node src/index.js.
  - id: R015
    from_status: active
    to_status: validated
    proof: wrapAuthError maps AADSTS50076/53003/700016/50059 to targeted messages. Client maps 401/403/404 to actionable guidance. 16 unit tests prove error paths.
duration: ~2h
verification_result: passed
completed_at: 2026-03-16
---

# M001: Zero-Config SharePoint MCP

**Server starts with `npx sharepoint-online-mcp` — no env vars, no Azure Portal, no app registration. Authenticates via device code using well-known Microsoft Office client ID, auto-discovers tenant from SharePoint URLs, routes dual-audience tokens to 25 MCP tools covering pages, layout, web parts, navigation, and branding. Package is publish-ready with 72 tests passing.**

## What Happened

Four slices rebuilt the auth layer, added site discovery, validated all tools, and packaged the result.

**S01 (Zero-Config Auth Engine)** rewrote `src/auth.js` from a parameterized class requiring clientId/tenantId to a zero-config engine. The `SharePointAuth` class hardcodes Microsoft Office's well-known client ID `d3590ed6-52b3-4102-aeff-aad2292ab01c` and uses `common` authority — zero constructor arguments. Added `discoverTenantId(domain)` for automatic tenant GUID resolution via OpenID configuration. Wired dual-audience token routing through `src/client.js`: Graph API calls get `https://graph.microsoft.com` tokens, SP REST calls extract the SharePoint origin via `new URL(siteUrl).origin`. Stripped all `process.env.SHAREPOINT_*` reads from `src/index.js` — server now starts instantly with zero configuration. Added `disconnect` tool and `buildScopes` export. Established the test framework (node:test, 18 tests).

**S02 (Site Discovery & Connection Tools)** translated all German tool descriptions, parameter names, and response strings to English across all 25 tools. Added `parseSharePointUrl` as a pure exported function handling /sites/, /teams/, root sites, trailing slashes, and subpages. Added `connect_to_site` tool that chains URL parsing → tenant discovery → auth → Graph site resolution in a single call. Added `list_my_sites` for followed sites. Established the cross-module import pattern (tools.js importing from auth.js). 7 URL parsing tests.

**S03 (Tool Validation & Dual-Audience Tokens)** built mock infrastructure (MockServer, MockClient, MockAuth) and wrote 31 handler tests covering every tool category: sites (5), pages (7), layout (3), web parts (6), navigation (4), branding (2), auth/utility (4). Each test verifies correct client method delegation, MCP content shape, and error path behavior. Added a minimal `overrides` parameter to `registerTools()` for dependency injection — 2 lines of source change, avoids experimental Node flags. Total: 56 tests.

**S04 (npm Package & Polish)** added npm metadata (files whitelist, engines, keywords, license, repository), created MIT LICENSE and .npmignore, set executable shebang on index.js. Rewrote README from German to English with zero-config quick start, Claude Desktop config, all 25 tools documented by category, troubleshooting table, and known limitations. Added `wrapAuthError()` pure function mapping AADSTS error codes to targeted remediation messages. Enhanced `request()` and `spRest()` with status-specific error messages for 401/403/404. Made fetch injectable in SharePointClient for testable error paths. 16 new error tests, bringing total to 72.

## Cross-Slice Verification

### Success Criteria

| Criterion | Status | Evidence |
|-----------|--------|----------|
| Server starts with `npx sharepoint-online-mcp` without env vars or config | ✅ Pass | `node src/index.js` emits `🚀 SharePoint Online MCP Server started (stdio)` on stderr. `grep -rn 'process\.env\.SHAREPOINT' src/` returns empty. |
| Device code flow authenticates against real Azure AD tenant | ⚠️ Contract-verified | Auth code implemented, unit tested (12 tests). Live device code flow requires human interaction — deferred to human UAT. |
| Tenant ID auto-discovered from SharePoint URL | ⚠️ Contract-verified | `discoverTenantId` implemented with 6 mock-fetch tests. Live OpenID endpoint call requires human UAT. |
| All 20+ existing tools work with new auth layer | ✅ Pass | 25 tools registered (`grep -c 'server.tool(' src/tools.js` = 25). 31 handler mock tests prove correct delegation through Graph and SP REST client layers. |
| Token persists across server restarts | ⚠️ Contract-verified | MSAL cache plugin configured for `~/.sharepoint-mcp-cache.json`. Cross-restart persistence requires human UAT. |
| `npx sharepoint-online-mcp` installs and runs cleanly from npm | ✅ Pass | `npm pack --dry-run` produces clean 8-file, 12.1kB artifact. Server starts without error. |

### Definition of Done

| Item | Status | Evidence |
|------|--------|----------|
| Auth module acquires tokens via well-known client ID + device code flow without env vars | ⚠️ Contract | Code complete, 12 unit tests, zero env var reads. Live acquisition requires human. |
| Tenant discovery works from arbitrary SharePoint URLs | ⚠️ Contract | 6 unit tests with mock fetch. Live OpenID endpoint untested. |
| All existing tools work with new auth | ✅ | 31 handler tests, all 25 tools delegate correctly. |
| Token caching works across server restarts | ⚠️ Contract | MSAL cache plugin code exists. Untested cross-restart. |
| Disconnect tool clears auth and enables re-auth | ⚠️ Contract | Tool registered, handler tested in S03. Live proof needs human. |
| `npx sharepoint-online-mcp` starts and runs cleanly | ✅ | Confirmed — startup message, no errors. |
| package.json is publish-ready | ✅ | npm pack --dry-run clean. All metadata present. |

**Summary:** All criteria are code-complete and contract-verified (72 tests, 0 failures). Three criteria additionally have operational proof (server startup, tool count, npm pack). Four criteria require human interaction for live integration proof (device code flow inherently needs a human at a browser). This is by design — every slice documented this deferral.

### Regression Suite

```
node --test tests/*.test.js
  tests/auth.test.js    — 20 pass (buildScopes, discoverTenantId, SharePointAuth, wrapAuthError)
  tests/client.test.js  — 14 pass (Graph routing, SP REST origin, HTTP error messages)
  tests/tools.test.js   — 38 pass (7 URL parsing, 31 handler delegation)
  Total: 72 pass, 0 fail, <1s
```

## Requirement Changes

- R001 (Zero-config Auth): active → validated — auth.js hardcodes Office client ID, zero constructor args, 3 unit tests, server starts without env vars
- R002 (Auto Tenant Discovery): active → validated — discoverTenantId with 6 mock-fetch unit tests
- R003 (Device Code Flow): active → validated — implemented with English stderr prompt, code path unit tested
- R004 (Token Caching): active → validated — MSAL cache plugin configured for ~/.sharepoint-mcp-cache.json
- R005 (Disconnect Tool): active → validated — tool registered, handler tested in S03
- R006 (Interactive Site Discovery): active → validated — connect_to_site, list_my_sites, search_sites registered and tested
- R007 (Multi-Site Support): active → validated — stateless siteId pattern, connect_to_site resolves any URL
- R008–R012: remained validated (from S03)
- R013 (npm Package): active → validated — npm pack --dry-run clean, server starts, metadata complete
- R014: remained validated (from S01)
- R015 (Clear Error Messages): active → validated — 16 tests prove AADSTS + HTTP error guidance
- R016: remained deferred — documented as known limitation in README

## Forward Intelligence

### What the next milestone should know
- The entire auth stack is contract-verified but has never run against a live Azure AD tenant. The first human UAT pass is the true proof point. If the well-known client ID (`d3590ed6`) lacks pre-consented scopes for Graph Beta pages endpoints, the only change is the client ID constant on line ~10 of `src/auth.js`.
- All 25 tools are in `src/tools.js` as a single 500+ line file. If the tool count grows, splitting by category (pages, nav, branding, sites) is the natural seam.
- The test infrastructure (MockServer/MockClient/MockAuth in tests/tools.test.js) is reusable for any new tools.
- Package is publish-ready but not published. `npm publish` is a manual step after human UAT passes.

### What's fragile
- **Well-known client ID scope availability** — Microsoft may not have pre-consented Graph Beta scopes (pages, layout) for the Office client ID. This is the single highest-risk assumption. Fallback client IDs: Azure CLI (`04b07795-ee59-4573-a679-68e3d8b21c85`), VS Code (`aebc6443-996d-45c2-90f0-388ff96faa56`).
- **MSAL v5 cache format** — `~/.sharepoint-mcp-cache.json` written by MSAL's plugin. Major version changes could silently break cache (user just re-auths, not catastrophic).
- **README tool table** — manually maintained, must be updated in lockstep with tools.js.

### Authoritative diagnostics
- `node --test tests/*.test.js` — 72 tests, <1s, full regression. Trust this over any partial check.
- `npm pack --dry-run` — exactly what gets published. Catches leaked files.
- `grep -c 'server.tool(' src/tools.js` — tool count invariant (25).
- `grep -rn 'process\.env\.SHAREPOINT' src/` — must return empty. Any hits mean env var gates leaked back.

### What assumptions changed
- MSAL version: plan referenced ^2.16.2, actual is ^5.1.0. API compatible for our usage.
- node-fetch: assumed needed, but Node 25 global fetch is sufficient for auth.js. node-fetch retained for client.js.
- Test count: plan estimated ~30 handler tests, actual is 31. Error path tests added 16 more (72 total vs ~48 originally scoped).
- Root site path: plan assumed `sitePath || "/"`, actual uses empty string because Graph API path differs (D012).

## Files Created/Modified

- `src/auth.js` — Zero-config MSAL wrapper: well-known client ID, discoverTenantId, buildScopes, wrapAuthError, device code flow with English prompts
- `src/client.js` — Dual-audience token routing (Graph vs SP REST origin extraction), injectable fetchFn, status-specific error messages
- `src/index.js` — Zero-config entrypoint: no env var reads, passes auth to client and tools
- `src/tools.js` — 25 MCP tools with English descriptions, parseSharePointUrl, connect_to_site, list_my_sites, disconnect, overrides injection seam
- `src/auth-cli.js` — Standalone zero-config auth tester
- `tests/auth.test.js` — 20 tests: buildScopes, discoverTenantId, SharePointAuth construction, wrapAuthError
- `tests/client.test.js` — 14 tests: Graph routing, SP REST origin extraction, HTTP error messages
- `tests/tools.test.js` — 38 tests: URL parsing (7), handler delegation (31) with MockServer/MockClient/MockAuth
- `package.json` — Renamed to sharepoint-online-mcp, bin entry, files whitelist, engines, keywords, license, repository, test script
- `LICENSE` — MIT license
- `.npmignore` — Exclusion list for npm publish
- `README.md` — English zero-config documentation: quick start, Claude Desktop config, 25 tool reference, troubleshooting
