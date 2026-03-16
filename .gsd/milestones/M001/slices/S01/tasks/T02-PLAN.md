---
estimated_steps: 6
estimated_files: 6
---

# T02: Wire dual-audience tokens through client and enable zero-config server startup

**Slice:** S01 — Zero-Config Auth Engine
**Milestone:** M001

## Description

Connect the refactored auth engine from T01 to the rest of the codebase. Update `client.js` to pass the correct resource URL for each API type (Graph vs SP REST), strip env var gates from `index.js` so the server starts immediately, add a `disconnect` MCP tool, update `auth-cli.js` for zero-config, and declare all runtime dependencies in `package.json`.

## Steps

1. **Update `src/client.js`** for dual-audience token routing:
   - In `request(url, options)`: change `await this.auth.getAccessToken()` → `await this.auth.getAccessToken("https://graph.microsoft.com")`. This method is only used by `graph()` and `graphBetaReq()` which always hit `graph.microsoft.com`.
   - In `spRest(siteUrl, apiPath, options)`: extract the hostname from `siteUrl` using `new URL(siteUrl).origin` (this gives `https://contoso.sharepoint.com`). Change `await this.auth.getAccessToken()` → `await this.auth.getAccessToken(origin)` where `origin` is the extracted value. This correctly handles `contoso.sharepoint.com`, `contoso-admin.sharepoint.com`, and `contoso-my.sharepoint.com`.
   - No other changes to client.js. All existing API methods stay the same.

2. **Update `src/index.js`** for zero-config startup:
   - Remove the `CLIENT_ID` and `TENANT_ID` env var reads (`process.env.SHAREPOINT_CLIENT_ID`, etc.)
   - Remove the entire `if (!CLIENT_ID)` error block (lines ~20-40)
   - Change `new SharePointAuth(CLIENT_ID, TENANT_ID)` → `new SharePointAuth()` (no args)
   - Remove the `TENANT_ID` variable
   - Update the server name from `"sharepoint-mcp"` to `"sharepoint-online-mcp"`
   - Update the server description to English: `"MCP Server for SharePoint Online — create, edit, and design SharePoint sites and pages via Claude. Zero-config: authenticates via device code flow, no Azure Portal setup required."`
   - Update the startup message to English: `"🚀 SharePoint Online MCP Server started (stdio)\n"`
   - Pass `auth` to `registerTools(server, client)` as well — it'll be needed for the disconnect tool: `registerTools(server, client, auth)`

3. **Add `disconnect` tool in `src/tools.js`**:
   - The `registerTools` function signature changes from `(server, client)` to `(server, client, auth)`
   - Add a new tool registration at the end of the function:
     ```
     server.tool("disconnect", "Disconnect from SharePoint and clear cached authentication. Use this to switch accounts or re-authenticate.", {}, async () => {
       await auth.logout();
       return { content: [{ type: "text", text: "Disconnected. Authentication cache cleared. You will be prompted to re-authenticate on the next request." }] };
     });
     ```
   - This is the only change to tools.js — all 20+ existing tools stay untouched.

4. **Update `src/auth-cli.js`** for zero-config:
   - Remove env var reads (`process.env.SHAREPOINT_CLIENT_ID`, etc.)
   - Construct `new SharePointAuth()` with no args
   - Update test call from `auth.getAccessToken()` → `auth.getAccessToken("https://graph.microsoft.com")`
   - Update messages to English

5. **Update `package.json`**:
   - Change `"name"` from `"sharepoint-mcp"` to `"sharepoint-online-mcp"`.
   - Change `"bin"` key from `"sharepoint-mcp"` to `"sharepoint-online-mcp"`.
   - Add `"dependencies"` block:
     ```json
     "dependencies": {
       "@azure/msal-node": "^2.16.2",
       "@modelcontextprotocol/sdk": "^1.6.1",
       "node-fetch": "^3.3.2",
       "zod": "^3.24.1"
     }
     ```
   - Verify the actual installed versions by checking `node_modules/*/package.json` or `package-lock.json` and use compatible version ranges. If `node_modules` doesn't exist, use the versions above as reasonable defaults.

6. **Write `tests/client.test.js`**:
   - Test that `request()` calls `auth.getAccessToken` with `"https://graph.microsoft.com"`:
     - Create a mock auth object with `getAccessToken` that records the argument passed
     - Create a `SharePointClient` with the mock auth
     - Mock global `fetch` to return a successful response
     - Call `client.graph("/me")` and assert `getAccessToken` was called with `"https://graph.microsoft.com"`
   - Test that `spRest()` calls `auth.getAccessToken` with the SP origin:
     - Same setup, call `client.spRest("https://contoso.sharepoint.com/sites/marketing", "web")` 
     - Assert `getAccessToken` was called with `"https://contoso.sharepoint.com"`
   - Test with `-admin` and `-my` SharePoint subdomains to verify hostname extraction handles all variants.

## Must-Haves

- [ ] `client.js` `request()` passes `"https://graph.microsoft.com"` resource to auth
- [ ] `client.js` `spRest()` extracts origin from `siteUrl` and passes it as resource to auth
- [ ] `index.js` has zero references to `process.env.SHAREPOINT_CLIENT_ID` or `SHAREPOINT_TENANT_ID`
- [ ] `index.js` constructs `SharePointAuth()` with no arguments
- [ ] `disconnect` tool is registered in `tools.js`
- [ ] `package.json` name is `sharepoint-online-mcp` and declares all runtime dependencies
- [ ] Server starts without crash: `node src/index.js` emits startup message on stderr
- [ ] All tests pass: `node --test tests/client.test.js`

## Verification

- `node --test tests/client.test.js` — all tests pass
- `timeout 3 node src/index.js 2>&1 || true` — stderr contains "Server started" or similar, no "SHAREPOINT_CLIENT_ID" error
- `grep -rn "process\.env\.SHAREPOINT" src/ | grep -v "\/\/"` — returns empty (no env var references outside comments)
- `node -e "import('./src/tools.js').then(m => console.log(typeof m.registerTools))"` — prints `function`

## Observability Impact

- Signals added/changed: Server startup message confirms zero-config mode. Disconnect tool provides explicit feedback on cache clear.
- How a future agent inspects this: `~/.sharepoint-mcp-cache.json` presence confirms auth happened. Server stderr output shows auth state.
- Failure state exposed: Auth errors now include the resource URL that failed, SP REST errors include the SharePoint hostname attempted.

## Inputs

- `src/auth.js` — refactored by T01, exports `SharePointAuth` (no-arg constructor) and `discoverTenantId`
- T01 established `node --test` as the test runner and created `tests/` directory

## Expected Output

- `src/client.js` — dual-audience token routing (Graph resource for `request()`, SP origin for `spRest()`)
- `src/index.js` — zero-config startup, no env var gates
- `src/tools.js` — `disconnect` tool added, signature includes `auth` param
- `src/auth-cli.js` — zero-config standalone auth test
- `tests/client.test.js` — unit tests for resource routing
- `package.json` — renamed, dependencies declared
