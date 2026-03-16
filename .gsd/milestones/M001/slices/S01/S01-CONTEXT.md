---
id: S01
milestone: M001
status: ready
---

# S01: Zero-Config Auth Engine — Context

<!-- Slice-scoped context. Milestone-only sections (acceptance criteria, completion class,
     milestone sequence) do not belong here — those live in the milestone context. -->

## Goal

Server starts without env vars, device code authenticates against a real tenant using well-known client ID, token is cached, Graph API call succeeds — proven against live Azure AD.

## Why this Slice

This is the foundation for everything else. Every downstream slice (site discovery, tool validation, npm packaging) depends on a working zero-config auth engine. It retires the two highest risks early: whether the well-known client ID has sufficient scopes, and whether dual-audience tokens work.

## Scope

### In Scope

- Replace constructor-injected `clientId` / `tenantId` with well-known Microsoft Office client ID (`d3590ed6-52b3-4102-aeff-aad2292ab01c`)
- Remove `SHAREPOINT_CLIENT_ID` / `SHAREPOINT_TENANT_ID` env var requirements from server startup
- Implement `discoverTenantId(sharepointDomain)` via OpenID Discovery endpoint
- Lazy auth trigger — server starts instantly, auth happens on first tool call that needs a token
- Device code flow returns structured MCP tool result (URL + user code) so Claude can present it conversationally; stderr as fallback only
- `getAccessToken(resource)` supports per-resource token acquisition (Graph vs SP REST audience)
- Token cache in `~/.sharepoint-online-mcp/cache.json` (dedicated directory)
- Single cache file holds tokens for multiple tenants (MSAL handles per-account keying)
- Silent token renewal from cache on subsequent calls
- Rename all internal references from `sharepoint-mcp` to `sharepoint-online-mcp` (package.json name + bin, MCP server name in index.js, cache path)
- All user-facing messages in English (replace current German strings)
- Clear, actionable error messages on auth failure — no auto-retry, return specific error + retry hint

### Out of Scope

- Site discovery / `connect_to_site` tool (S02)
- Disconnect/logout tool (S02 — needs site context concept)
- Validation of all 20+ existing tools with new auth (S03)
- npm packaging, README rewrite (S04)
- Alternative auth flows (browser redirect, auth code)
- Conditional Access workarounds
- TypeScript migration
- Admin-level features

## Constraints

- Must use well-known first-party client ID — no custom Azure AD app registration
- Must work over MCP stdio (device code cannot use interactive terminal prompts — deliver via tool result)
- No build step — plain ESM JavaScript, npx-ready
- Node.js >= 18
- Refactor existing code, do not rewrite from scratch (D004)
- Package name is `sharepoint-online-mcp` everywhere — npm name, bin command, MCP server name, cache directory (D005)

## Integration Points

### Consumes

- `@azure/msal-node` — `PublicClientApplication`, device code flow, token cache serialization
- Microsoft OpenID Discovery endpoint — `https://login.microsoftonline.com/{domain}/.well-known/openid-configuration` for tenant ID resolution
- Existing `src/auth.js` (95 lines) — refactored, not replaced
- Existing `src/index.js` (71 lines) — env var gating removed, lazy init added
- Existing `src/client.js` (319 lines) — `getAccessToken()` call signature changes to accept resource parameter

### Produces

- `SharePointAuth` class with `getAccessToken(resource)` — handles well-known client ID, device code flow, per-resource token caching, and silent renewal
- `discoverTenantId(sharepointDomain)` function — resolves tenant GUID from SharePoint domain
- Server starts without env vars — no startup crash, auth triggered lazily
- Token cache at `~/.sharepoint-online-mcp/cache.json` — persists across restarts, multi-tenant capable
- Proof that well-known client ID token can call Graph Beta sites/pages endpoint (risk retirement)

## Open Questions

- Will `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Office client) get `Sites.ReadWrite.All` scope via `.default`? — Current thinking: test during execution; if not, fall back to Azure CLI client ID `04b07795-8ddb-461a-bbee-02f9e1bf7b46`
- Does SP REST API accept tokens acquired with Graph resource scope, or does it need a separate token with `https://{tenant}.sharepoint.com` audience? — Current thinking: likely needs separate token acquisition; `getAccessToken(resource)` is designed for this
- Disconnect clears only current tenant's tokens (not all) — implementation detail: use MSAL's `removeAccount()` for the active account rather than deleting the whole cache file
