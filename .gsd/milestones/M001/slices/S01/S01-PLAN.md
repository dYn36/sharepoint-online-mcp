# S01: Zero-Config Auth Engine

**Goal:** Server starts without env vars, device code authenticates against a real tenant using well-known client ID, token is cached, Graph API call succeeds — proven against live Azure AD
**Demo:** `node src/index.js` starts without env vars or config. On first tool call, device code prompt appears on stderr. After auth, Graph API and SP REST calls succeed with correct per-resource tokens. Token persists across restarts.

## Must-Haves

- Auth uses Microsoft Office well-known client ID `d3590ed6-52b3-4102-aeff-aad2292ab01c` — no constructor params, no env vars
- `discoverTenantId(domain)` resolves tenant GUID from SharePoint domain via OpenID config endpoint
- `getAccessToken(resource)` acquires per-resource tokens: `https://graph.microsoft.com/.default` for Graph, `https://{host}.sharepoint.com/.default` for SP REST
- Device code flow fires only on first token request; subsequent calls use `acquireTokenSilent`
- Token cache persists to `~/.sharepoint-mcp-cache.json` across restarts
- Server starts immediately with `node src/index.js` — no env var checks
- `disconnect` MCP tool clears token cache and enables re-auth
- `package.json` declares all runtime dependencies

## Proof Level

- This slice proves: integration (auth engine acquires real tokens, but full tool validation is S03)
- Real runtime required: yes (manual test against live Azure AD tenant)
- Human/UAT required: yes (device code flow requires browser interaction)

## Verification

- `node --test tests/auth.test.js` — unit tests pass for tenant discovery URL parsing, scope construction, auth class construction
- `node --test tests/client.test.js` — unit tests pass for dual-audience token routing in client
- `node src/index.js 2>&1 &; PID=$!; sleep 2; kill $PID 2>/dev/null; wait $PID 2>/dev/null` — server starts without crash, no env var error on stderr
- `grep -rn "process\.env\.SHAREPOINT" src/ | grep -v "^.*:.*\/\/"` — zero hits (no env var references outside comments)
- Manual: device code flow → token → Graph API call succeeds (human verification, not automated)

## Observability / Diagnostics

- Runtime signals: stderr messages for auth state transitions — device code prompt, silent token renewal, auth failure with error detail
- Inspection surfaces: `~/.sharepoint-mcp-cache.json` existence and content (token cache), stderr log output
- Failure visibility: Auth errors include resource URL, HTTP status, and MSAL error code/message. Tenant discovery errors include the domain attempted and HTTP response.
- Redaction constraints: Never log access tokens, refresh tokens, or user codes to any persistent log. Stderr prompt for device code is transient and expected.

## Integration Closure

- Upstream surfaces consumed: none (first slice)
- New wiring introduced in this slice: `SharePointAuth` zero-config class used by `SharePointClient`, `disconnect` tool registered in MCP server, `package.json` dependency declarations
- What remains before the milestone is truly usable end-to-end: S02 (site discovery tools), S03 (validate all 20+ tools with dual-audience tokens), S04 (npm packaging and docs)

## Tasks

- [ ] **T01: Refactor auth.js to zero-config multi-resource auth engine with unit tests** `est:1h`
  - Why: This is the core risk-retiring task. Proves well-known client ID pattern works, tenant discovery parses correctly, per-resource scope construction is correct. Unblocks all downstream work.
  - Files: `src/auth.js`, `tests/auth.test.js`, `package.json`
  - Do: (1) Add `test` script to package.json using `node --test`. (2) Rewrite `auth.js`: remove constructor params, hardcode Office client ID, add `discoverTenantId(domain)` that calls OpenID config endpoint and parses tenant GUID from `token_endpoint`, change `getAccessToken(scopes)` → `getAccessToken(resource)` that builds `["{resource}/.default"]` scopes, use `common` authority. (3) Write `tests/auth.test.js` with unit tests: tenant discovery URL construction and GUID parsing (mock fetch), scope construction for Graph and SP resources, auth class instantiates without args, error handling for failed tenant discovery.
  - Verify: `node --test tests/auth.test.js` — all tests pass
  - Done when: `auth.js` exports zero-config `SharePointAuth` class and `discoverTenantId` function, all unit tests pass

- [ ] **T02: Wire dual-audience tokens through client and enable zero-config server startup** `est:45m`
  - Why: Completes the slice by connecting the new auth engine to the rest of the codebase. Without this, auth.js is refactored but the server still crashes on missing env vars.
  - Files: `src/client.js`, `src/index.js`, `src/tools.js`, `src/auth-cli.js`, `tests/client.test.js`, `package.json`
  - Do: (1) Update `client.js`: `request()` calls `this.auth.getAccessToken("https://graph.microsoft.com")`, `spRest()` extracts hostname from `siteUrl` param and calls `this.auth.getAccessToken("https://{hostname}")`. (2) Update `index.js`: remove `CLIENT_ID`/`TENANT_ID` env var reads and the error block, construct `new SharePointAuth()` with no args, update server description to English. (3) Add `disconnect` tool in `tools.js`: calls `auth.logout()`, returns confirmation. (4) Update `auth-cli.js` for zero-config. (5) Update `package.json`: name → `sharepoint-online-mcp`, add `dependencies` block with `@azure/msal-node`, `@modelcontextprotocol/sdk`, `node-fetch`, `zod`. (6) Write `tests/client.test.js`: verify `request()` passes Graph resource, `spRest()` passes SP resource (mock auth).
  - Verify: `node --test tests/client.test.js` — all tests pass. `node src/index.js` starts without crash (stderr shows server started, no env var error). `grep -rn "process\.env\.SHAREPOINT" src/ | grep -v "\/\/"` returns empty.
  - Done when: Server starts with zero env vars, client routes correct resource per API type, disconnect tool is registered, package.json is complete with all deps

## Files Likely Touched

- `src/auth.js`
- `src/client.js`
- `src/index.js`
- `src/tools.js`
- `src/auth-cli.js`
- `tests/auth.test.js`
- `tests/client.test.js`
- `package.json`
