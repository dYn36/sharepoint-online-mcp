# S01: Zero-Config Auth Engine — UAT

**Milestone:** M001
**Written:** 2026-03-16

## UAT Type

- UAT mode: mixed (artifact-driven for code structure, live-runtime for auth flow)
- Why this mode is sufficient: Auth engine correctness is split — URL parsing and scope construction are verifiable via unit tests, but device code flow and token acquisition require a real Azure AD tenant with browser interaction.

## Preconditions

- Node.js 18+ installed (Node 25 preferred — built-in fetch used)
- `npm install` completed in project root (node_modules present)
- Access to an Azure AD tenant with at least one SharePoint Online site
- A web browser available for device code authentication
- No `SHAREPOINT_CLIENT_ID` or `SHAREPOINT_TENANT_ID` env vars set (verify with `echo $SHAREPOINT_CLIENT_ID`)

## Smoke Test

Run `node src/index.js` — server should emit `🚀 SharePoint Online MCP Server started (stdio)` on stderr within 1 second. No env var errors, no crash.

## Test Cases

### 1. Zero-config server startup

1. Unset any SharePoint env vars: `unset SHAREPOINT_CLIENT_ID SHAREPOINT_TENANT_ID`
2. Run `node src/index.js`
3. **Expected:** stderr shows `🚀 SharePoint Online MCP Server started (stdio)` immediately. No error messages. Process stays running (listening on stdio).
4. Kill the process with Ctrl+C.

### 2. Unit tests pass

1. Run `node --test tests/auth.test.js`
2. **Expected:** 12/12 tests pass — buildScopes (3), discoverTenantId (6), SharePointAuth (3)
3. Run `node --test tests/client.test.js`
4. **Expected:** 6/6 tests pass — Graph routing (2), SP REST origin extraction (3), path stripping (1)

### 3. No env var gates in source

1. Run `grep -rn "process\.env\.SHAREPOINT" src/ | grep -v "\/\/"`
2. **Expected:** Zero output. No env var references outside comments.
3. Run `grep -rn "process\.env\.CLIENT_ID\|process\.env\.TENANT_ID" src/ | grep -v "\/\/"`
4. **Expected:** Zero output.

### 4. Device code flow authentication (live Azure AD)

1. Start the server: `node src/index.js 2>/tmp/sp-mcp-stderr.log &`
2. Send a tool call via stdin that triggers auth (e.g., a search_sites call via MCP JSON-RPC)
3. Check stderr: `cat /tmp/sp-mcp-stderr.log`
4. **Expected:** stderr shows a device code message with a URL (https://microsoft.com/devicelogin) and a user code
5. Open the URL in a browser, enter the code, sign in with your Azure AD account
6. **Expected:** After successful browser auth, the tool call completes (or at least progresses past auth)

### 5. Token caching across restarts

1. After completing test case 4 (authenticated), kill the server process
2. Verify cache exists: `ls -la ~/.sharepoint-mcp-cache.json`
3. **Expected:** File exists, non-empty, contains JSON
4. Restart the server: `node src/index.js 2>/tmp/sp-mcp-stderr2.log &`
5. Send the same tool call as test case 4
6. Check stderr: `cat /tmp/sp-mcp-stderr2.log`
7. **Expected:** No device code prompt appears — token acquired silently from cache

### 6. Disconnect tool clears auth

1. With an authenticated session running, send a `disconnect` tool call via MCP
2. **Expected:** Tool returns confirmation message that auth was cleared
3. Verify cache is cleared: `cat ~/.sharepoint-mcp-cache.json`
4. **Expected:** Cache file is empty or deleted
5. Send a tool call that requires auth
6. **Expected:** Device code prompt reappears on stderr — fresh authentication required

### 7. Standalone auth CLI tester

1. Delete cache if it exists: `rm -f ~/.sharepoint-mcp-cache.json`
2. Run `node src/auth-cli.js`
3. **Expected:** Device code prompt appears on stderr with verification URL and user code
4. Complete browser auth
5. **Expected:** CLI outputs a success message and a Graph API access token (or confirmation of token acquisition)

### 8. Dual-audience token routing (live)

1. With an authenticated session, send a tool call that uses Graph API (e.g., search_sites)
2. **Expected:** Call succeeds — token was acquired for `https://graph.microsoft.com`
3. Send a tool call that uses SP REST API (e.g., get_navigation for a specific site URL)
4. **Expected:** Call succeeds — token was acquired for `https://{tenant}.sharepoint.com`
5. Both calls should work without re-authentication (tokens cached per-resource)

## Edge Cases

### Tenant discovery with non-standard SharePoint URLs

1. Call `discoverTenantId` (via auth-cli or test) with a `-admin` subdomain: e.g., `contoso-admin.sharepoint.com`
2. **Expected:** Tenant GUID is correctly resolved (the OpenID endpoint works with any valid domain)

### Invalid SharePoint domain

1. Attempt tenant discovery with a non-existent domain: `fakecorp12345.sharepoint.com`
2. **Expected:** Error message includes the domain attempted and a clear failure reason (HTTP error or network failure)

### Server startup with stale cache

1. Corrupt the cache file: `echo "invalid json" > ~/.sharepoint-mcp-cache.json`
2. Run `node src/index.js`
3. **Expected:** Server starts normally. On first tool call, device code flow triggers (stale/corrupt cache is ignored gracefully).

## Failure Signals

- Server crashes on startup with "SHAREPOINT_CLIENT_ID is required" or similar → env var gates not fully removed
- `grep` finds `process.env.SHAREPOINT` references in `src/` → incomplete refactoring
- Unit tests fail → auth logic regression
- Device code prompt never appears on tool call → auth not triggering lazily
- Token acquired but Graph API returns 401 → well-known client ID may lack required scopes (key risk for S03)
- SP REST call returns 401 after Graph call works → dual-audience routing broken or SP REST scopes not available on well-known client ID
- `disconnect` doesn't trigger re-auth → cache not properly cleared

## Requirements Proved By This UAT

- R001 (Zero-config Auth) — test cases 1, 4: server starts with no env vars, auth uses well-known client ID
- R002 (Auto Tenant Discovery) — test case 4 + edge cases: tenant resolved from SharePoint domain
- R003 (Device Code Flow) — test case 4: user authenticates via browser code entry
- R004 (Token Caching) — test case 5: tokens persist across server restarts
- R005 (Disconnect Tool) — test case 6: cache cleared, re-auth works
- R014 (No Hardcoded Secrets) — test case 3: grep confirms no secrets in source

## Not Proven By This UAT

- R006-R007 (Site Discovery) — S02 scope, no site discovery tools exist yet
- R008-R012 (Page/Layout/WebPart/Nav/Branding tools) — S03 scope, tools exist but not validated with new auth
- R013 (npm package) — S04 scope, npx invocation not tested
- R015 (Error messages) — partially covered by edge cases, full review in S04
- Whether the well-known client ID has sufficient scopes for all Graph Beta endpoints (pages, layout) — this is the key S03 risk

## Notes for Tester

- **Device code flow is interactive** — you must open a browser and enter the code. There's no way to automate this.
- **Test cases 4-8 require a real Azure AD tenant** — use any Microsoft 365 tenant where you have a user account with SharePoint access.
- **If your tenant blocks device code flow** (Conditional Access), all live tests will fail with a clear error. This is a known limitation (R016, deferred). Try a different tenant or a personal Microsoft account.
- **Token cache location** is `~/.sharepoint-mcp-cache.json` — this is in your home directory, not the project.
- **The `disconnect` tool is called via MCP protocol** — you'll need an MCP client (like Claude Desktop) or send raw JSON-RPC over stdio to invoke it.
- **Test case 8** (dual-audience) is the most important live test — it proves the core architectural bet that one well-known client ID can acquire tokens for both Graph and SharePoint REST APIs.
