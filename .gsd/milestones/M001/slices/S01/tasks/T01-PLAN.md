---
estimated_steps: 5
estimated_files: 3
---

# T01: Refactor auth.js to zero-config multi-resource auth engine with unit tests

**Slice:** S01 — Zero-Config Auth Engine
**Milestone:** M001

## Description

Rewrite `src/auth.js` from a parameterized auth class (requires `clientId`/`tenantId` constructor args) to a zero-config multi-resource auth engine. This is the highest-risk task in the slice — it proves the well-known client ID pattern, tenant discovery, and per-resource token acquisition all work structurally. Also establishes the test framework (Node's built-in `node:test`) since the project has none.

## Steps

1. **Add test infrastructure to `package.json`**: Add a `"test"` script: `"node --test"`. This uses Node's built-in test runner — no dependencies needed, matches the no-build-step constraint.

2. **Rewrite `src/auth.js`**:
   - Remove `clientId` and `tenantId` constructor parameters entirely.
   - Hardcode the Microsoft Office well-known client ID: `d3590ed6-52b3-4102-aeff-aad2292ab01c`.
   - Use `common` as the MSAL authority (allows any Azure AD tenant to authenticate).
   - Add exported `async function discoverTenantId(domain)`:
     - Takes a SharePoint domain like `contoso.sharepoint.com`
     - Fetches `https://login.microsoftonline.com/{domain}/.well-known/openid-configuration`
     - Parses the `token_endpoint` field from the JSON response
     - Extracts the tenant GUID from the URL path (format: `https://login.microsoftonline.com/{tenantGuid}/oauth2/v2.0/token`)
     - Returns the tenant GUID string
     - Throws a descriptive error if fetch fails or response format is unexpected
   - Change `getAccessToken(scopes = DEFAULT_SCOPES)` → `getAccessToken(resource)`:
     - `resource` is a full URL like `https://graph.microsoft.com` or `https://contoso.sharepoint.com`
     - Constructs scopes as `[`${resource}/.default`]`
     - Remove the `DEFAULT_SCOPES` constant entirely
     - Silent acquisition and device code flow logic stay the same, just with new scope format
   - Update device code callback message to English (current is German)
   - Keep the existing `logout()` method and cache plugin pattern unchanged
   - Keep `TOKEN_CACHE_PATH` unchanged (`~/.sharepoint-mcp-cache.json`)

3. **Write `tests/auth.test.js`** using `node:test` and `node:assert`:
   - **Test: `discoverTenantId` constructs correct URL** — Mock `global.fetch` (or the imported fetch), call `discoverTenantId("contoso.sharepoint.com")`, assert it fetched `https://login.microsoftonline.com/contoso.sharepoint.com/.well-known/openid-configuration`.
   - **Test: `discoverTenantId` parses tenant GUID from token_endpoint** — Mock fetch to return `{ token_endpoint: "https://login.microsoftonline.com/72f988bf-86f1-41af-91ab-2d7cd011db47/oauth2/v2.0/token" }`, assert return value is `"72f988bf-86f1-41af-91ab-2d7cd011db47"`.
   - **Test: `discoverTenantId` throws on fetch failure** — Mock fetch to reject or return non-OK, assert the function throws with a descriptive message including the domain attempted.
   - **Test: `getAccessToken` constructs correct scopes for Graph** — Cannot easily test without MSAL mock, but verify that the `SharePointAuth` class can be constructed without arguments (no throw). This is a structural smoke test.
   - **Test: scope construction helper** — If you extract scope construction to a helper function (recommended), test that `buildScopes("https://graph.microsoft.com")` returns `["https://graph.microsoft.com/.default"]` and `buildScopes("https://contoso.sharepoint.com")` returns `["https://contoso.sharepoint.com/.default"]`.

   Note: `auth.js` imports `node-fetch`. In the test file, mock `fetch` before importing auth functions. Use `node:test`'s `mock` API or replace the global. Since auth.js uses `import fetch from 'node-fetch'`, the cleanest approach is to make `discoverTenantId` accept an optional `fetchFn` parameter for testability, defaulting to the imported `fetch`.

4. **Verify**: Run `node --test tests/auth.test.js` and confirm all tests pass.

5. **Sanity check**: Run `node -e "import('./src/auth.js').then(m => { console.log(typeof m.SharePointAuth, typeof m.discoverTenantId) })"` to confirm both exports are available.

## Must-Haves

- [ ] `SharePointAuth` constructor takes zero arguments
- [ ] Well-known client ID `d3590ed6-52b3-4102-aeff-aad2292ab01c` is hardcoded in auth.js
- [ ] `discoverTenantId(domain)` is exported and calls OpenID config endpoint
- [ ] `getAccessToken(resource)` constructs `["{resource}/.default"]` scopes
- [ ] `DEFAULT_SCOPES` array is removed (no more named Graph scopes)
- [ ] Device code callback message is in English
- [ ] Tests pass: `node --test tests/auth.test.js`

## Verification

- `node --test tests/auth.test.js` — all tests pass
- `node -e "import('./src/auth.js').then(m => { console.log(typeof m.SharePointAuth, typeof m.discoverTenantId) })"` — prints `function function`

## Inputs

- `src/auth.js` — current 95-line auth module, the starting point for refactor
- `package.json` — needs test script added

## Expected Output

- `src/auth.js` — refactored zero-config auth module with `discoverTenantId` and resource-scoped `getAccessToken`
- `tests/auth.test.js` — unit tests for tenant discovery, scope construction, class instantiation
- `package.json` — `test` script added
