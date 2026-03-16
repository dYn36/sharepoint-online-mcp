# S01: Zero-Config Auth Engine — Research

**Date:** 2026-03-16
**Depth:** Deep (unfamiliar integration: well-known client IDs, dual-audience tokens, MSAL scope model)

## Summary

The existing auth module (`src/auth.js`, 95 lines) already implements MSAL Device Code Flow with token caching — the structure is sound. The refactor is surgical: replace the constructor's `clientId`/`tenantId` params with hardcoded well-known client ID, add a `discoverTenantId(domain)` function, and — critically — change `getAccessToken()` to accept a `resource` parameter for dual-audience token acquisition.

The biggest design insight from research: **MSAL tokens are per-resource**. A token acquired for `https://graph.microsoft.com` cannot be used against `https://{tenant}.sharepoint.com` and vice versa. The existing `client.js` uses a single `getAccessToken()` call for both Graph API and SP REST API — this is the core breaking change. The auth class must acquire separate tokens per audience, and the client must pass the correct resource scope for each request type. MSAL's cache handles this transparently — a single device code authentication yields a refresh token that can silently acquire tokens for any resource the client ID has access to.

The well-known Office client ID (`d3590ed6-52b3-4102-aeff-aad2292ab01c`) is confirmed to work with device code flow and has pre-consented permissions for both Graph API and SharePoint. Tenant discovery via OpenID configuration endpoint is a simple unauthenticated HTTP call.

## Recommendation

Refactor `auth.js` to be a zero-config, multi-resource auth engine:

1. **Hardcode the Office well-known client ID** — no constructor params needed
2. **Add `discoverTenantId(sharepointDomain)`** — extract tenant name from SP domain, call OpenID config endpoint, parse tenant GUID from `token_endpoint`
3. **Change `getAccessToken(scopes)` → `getAccessToken(resource)`** — resource is either `https://graph.microsoft.com` or `https://{tenant}.sharepoint.com`. Scopes become `["{resource}/.default"]`. Device code flow fires only on first call; subsequent calls (even different resource) use `acquireTokenSilent` with the cached refresh token.
4. **Update `client.js`** — `request()` and `spRest()` must pass the correct resource to `getAccessToken()`. This means `spRest()` needs to know the SharePoint domain to construct the resource URL.
5. **Update `index.js`** — remove env var checks, remove `SharePointAuth` constructor args, start server immediately
6. **Add `disconnect` tool** — calls `auth.logout()`, already exists as method, just needs MCP tool registration
7. **Add `dependencies`** to `package.json` — `@azure/msal-node`, `@modelcontextprotocol/sdk`, `node-fetch`, `zod`

Build order: auth.js first (riskiest — proves well-known client ID + dual tokens work), then client.js (plumbing), then index.js (remove env vars), then disconnect tool (trivial).

## Implementation Landscape

### Key Files

- `src/auth.js` (95 lines) — MSAL Device Code Flow + cache plugin. Currently takes `clientId`/`tenantId` in constructor. Needs: hardcoded well-known client ID, `discoverTenantId()` function, resource-scoped `getAccessToken(resource)`, lazy tenant resolution. The `beforeCacheAccess`/`afterCacheAccess` cache plugin pattern is already correct for MSAL Node.
- `src/client.js` (225 lines) — Graph + SP REST wrapper. `request()` calls `this.auth.getAccessToken()` with no resource param. `spRest()` also calls the same method. Needs: `request()` → pass `https://graph.microsoft.com` resource, `spRest()` → pass `https://{tenant}.sharepoint.com` resource. The SP resource URL requires knowing the SharePoint domain, so `spRest()` already receives `siteUrl` — extract the hostname from it.
- `src/index.js` (55 lines) — Entrypoint. Currently reads `SHAREPOINT_CLIENT_ID` env var, crashes if missing. Needs: remove all env var checks, construct `SharePointAuth()` with no args, start server immediately.
- `src/tools.js` (530 lines) — 20+ MCP tools. No changes needed for S01 except adding a `disconnect` tool. Tools already pass `siteUrl` for SP REST calls and `siteId` for Graph calls — the plumbing is correct.
- `src/auth-cli.js` (30 lines) — Standalone auth test. Needs: remove env var requirements, use zero-config auth.
- `package.json` — Missing `dependencies` block entirely. Needs all runtime deps declared.

### Build Order

1. **`auth.js` refactor** — This is the riskiest piece and unblocks everything. Prove: (a) well-known client ID works with `acquireTokenByDeviceCode`, (b) `.default` scope returns useful permissions, (c) `acquireTokenSilent` with different resource scope works after initial auth, (d) tenant discovery from SharePoint domain works. This single file change retires the two highest risks from the roadmap.

2. **`client.js` dual-audience plumbing** — Update `request()` and `spRest()` to pass correct resource to auth. `graph()` and `graphBetaReq()` use `https://graph.microsoft.com/.default`. `spRest()` extracts hostname from `siteUrl` param and uses `https://{hostname}/.default`. Straightforward once auth.js is done.

3. **`index.js` zero-config entrypoint** — Remove env var gate, construct auth with no args. Trivial.

4. **`package.json` dependencies** — Declare all runtime deps. Verify `npm install` works.

5. **`disconnect` tool + `auth-cli.js` update** — Register disconnect MCP tool in tools.js, update auth-cli for zero-config. Low risk.

### Verification Approach

- **Unit: Tenant discovery URL parsing** — Test that `discoverTenantId("contoso.sharepoint.com")` correctly constructs the OpenID config URL and parses tenant GUID from the response. Can mock the HTTP call.
- **Unit: Resource scope construction** — Test that `getAccessToken("https://graph.microsoft.com")` uses scopes `["https://graph.microsoft.com/.default"]` and `getAccessToken("https://contoso.sharepoint.com")` uses `["https://contoso.sharepoint.com/.default"]`.
- **Integration: Server starts without env vars** — `node src/index.js` must not crash, should output "server started" on stderr.
- **Integration: Auth flow** — Full device code → token acquisition → Graph API call cycle. Requires real Azure AD tenant (manual verification).
- **Structural: No hardcoded secrets** — `grep -r "SHAREPOINT_CLIENT_ID\|SHAREPOINT_TENANT_ID" src/` should return zero hits after refactor (except maybe as legacy comment).

## Don't Hand-Roll

| Problem | Existing Solution | Why Use It |
|---------|------------------|------------|
| OAuth token acquisition + caching + refresh | `@azure/msal-node` `PublicClientApplication` | Already in use. Handles silent renewal, refresh tokens, cache serialization. Don't reimplement. |
| Tenant ID discovery | OpenID Configuration endpoint (`login.microsoftonline.com/{domain}/.well-known/openid-configuration`) | Unauthenticated, returns JSON with tenant GUID in `token_endpoint`. Single `fetch` call. |
| MCP tool registration | `@modelcontextprotocol/sdk` | Already in use for all 20+ tools. |

## Constraints

- **MSAL tokens are per-resource-per-scope** — Cannot get a single token valid for both Graph API (`aud: https://graph.microsoft.com`) and SP REST API (`aud: https://{tenant}.sharepoint.com`). Must make separate `acquireTokenSilent` calls per resource. MSAL cache handles this — one device code auth yields a refresh token usable for any resource.
- **`.default` scope only with well-known client IDs** — Cannot request specific scopes like `Sites.ReadWrite.All` individually; must use `{resource}/.default` which returns whatever is pre-consented for the client ID. Cannot mix `.default` with named scopes in same request.
- **No build step** — Plain ESM JavaScript. All deps must be runtime-importable without transpilation.
- **MCP stdio** — Device code prompt must go to stderr (already does). Cannot use interactive terminal prompts.
- **Tenant discovery requires SharePoint domain** — `discoverTenantId` needs a SharePoint hostname (e.g., `contoso.sharepoint.com`). This means tenant discovery happens lazily when a tool first provides a SharePoint URL, not at server startup. The authority URL (`login.microsoftonline.com/{tenantId}`) must be set before token acquisition, which means MSAL's `PublicClientApplication` may need to be re-instantiated or use `common` authority initially.

## Common Pitfalls

- **MSAL PCA authority is set at construction** — If using `common` as authority, device code flow works but the token may not have the right audience for SP REST. Solution: use `common` for the initial device code flow (Graph token), then once tenant is known, create a tenant-specific PCA for SP REST tokens. Or simpler: use `common` for all — MSAL resolves the actual tenant from the user's login. Test this.
- **Cache deserialization with multiple PCA instances** — If creating separate PCA instances for Graph vs SP, they must share the same cache file. MSAL's cache plugin API supports this but the `beforeCacheAccess`/`afterCacheAccess` must handle concurrent reads correctly.
- **SharePoint domain extraction from site URL** — URLs like `https://contoso.sharepoint.com/sites/marketing` are straightforward, but `https://contoso-admin.sharepoint.com` or `https://contoso-my.sharepoint.com` also exist. The resource scope should use the exact hostname from the URL, not try to normalize it.
- **`node-fetch` vs global `fetch`** — Code currently imports `node-fetch`. Node 18+ has global fetch, but `node-fetch` returns a different Response type. Keep `node-fetch` for now per Decision 006 (no build step, compatibility).

## Open Risks

- **Well-known client ID scope coverage** — The Office client ID (`d3590ed6-52b3-4102-aeff-aad2292ab01c`) should have `Sites.ReadWrite.All` and `Sites.Manage.All` pre-consented via `.default`, but this is unconfirmed until tested against a real tenant. If `.default` returns insufficient scopes, fallback is Azure CLI client ID (`04b07795-8ddb-461a-bbee-02f9e1bf7b46`). This risk is retired by manual integration testing in this slice.
- **`common` authority + SP REST resource** — Using `common` authority with `https://{tenant}.sharepoint.com/.default` scope should work (MSAL resolves tenant from login), but untested. If it fails, we need to discover tenant ID first and set tenant-specific authority.
- **Some tenants block device code flow** — Conditional Access policies can block this entirely. Documented as known limitation (R016 deferred).

## Skills Discovered

| Technology | Skill | Status |
|------------|-------|--------|
| Microsoft Graph API | `markpitt/claude-skills@microsoft-graph` | available (98 installs) |
| MSAL / Azure Auth | none found | — |
| SharePoint Online | none found | — |

## Sources

- MSAL tokens are per-resource-per-scope — separate acquireTokenSilent per API (source: [MSAL.js Resources and Scopes](https://learn.microsoft.com/en-us/entra/msal/javascript/browser/resources-and-scopes))
- Well-known Office client ID confirmed working with device code flow for Graph API (source: [GuardSix - How OAuth and Device Code Flows Get Abused](https://guardsix.com/blog/emerging-threats/how-oauth-and-device-code-flows-get-abused))
- Tenant ID discoverable from domain via OpenID configuration endpoint (source: [SharePoint Diary - How to Get Tenant ID](https://www.sharepointdiary.com/2019/04/how-to-get-tenant-id-in-sharepoint-online.html))
- `.default` scope returns all pre-consented permissions for the client (source: [Microsoft - Scopes and permissions](https://learn.microsoft.com/en-us/entra/identity-platform/scopes-oidc))
- MSAL Node device code flow API and cache plugin pattern (source: [MSAL Node - Acquiring tokens](https://learn.microsoft.com/en-us/entra/msal/javascript/node/acquire-token-requests))
- Dual Graph + SharePoint consent requires separate token acquisitions (source: [Cameron Dwyer - MSAL Graph and SharePoint Consent](https://camerondwyer.com/2022/03/11/how-to-combine-graph-sharepoint-permission-consent-into-a-single-msal-dialog-on-first-use/))
