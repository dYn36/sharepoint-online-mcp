# Decisions Register

<!-- Append-only. Never edit or remove existing rows.
     To reverse a decision, add a new row that supersedes it.
     Read this file at the start of any planning or research phase. -->

| # | When | Scope | Decision | Choice | Rationale | Revisable? |
|---|------|-------|----------|--------|-----------|------------|
| 001 | 2026-03-16 | auth | Which client ID to use for zero-config auth | Microsoft Office first-party client ID `d3590ed6-52b3-4102-aeff-aad2292ab01c` | Pre-registered in every Azure AD tenant, supports Device Code Flow, has Graph API access. No app registration needed. | Yes — can switch to Azure CLI client ID if scope issues arise |
| 002 | 2026-03-16 | auth | How to discover tenant ID | Auto-derive from SharePoint domain via OpenID Discovery endpoint (`login.microsoftonline.com/{domain}/.well-known/openid-configuration`) | Users shouldn't need to know their tenant ID. SharePoint URL is always known. | No |
| 003 | 2026-03-16 | auth | When to trigger authentication | Lazy — on first tool call that needs a token, not at server startup | Zero-config means no env vars checked at startup. Server starts instantly; auth happens when needed. | Yes |
| 004 | 2026-03-16 | arch | Keep existing code or rewrite | Refactor existing codebase | 20+ tools already implemented and well-structured. Auth layer swap is surgical. | No |
| 005 | 2026-03-16 | dist | Package name | `sharepoint-online-mcp` | User specified. Distinguishes from on-prem SharePoint. | No |
| 006 | 2026-03-16 | arch | Language / build step | Plain JavaScript ESM, no TypeScript, no build step | Matches existing codebase. npx-ready without compilation. Simpler distribution. | Yes |
| 007 | 2026-03-16 | auth | Token scope strategy | Request `https://{tenant}.sharepoint.com/.default` for SP REST and `https://graph.microsoft.com/.default` for Graph | Well-known client IDs may not support fine-grained scope requests. `.default` requests whatever scopes are pre-consented for the client. | Yes — may need to test specific scope lists |
| 008 | 2026-03-16 | auth | Fetch implementation for discoverTenantId | Node built-in global fetch (not node-fetch) | Node 25 has stable global fetch. Fewer deps, injectable fetchFn param for testing. | No |
| 009 | 2026-03-16 | test | Test framework | node:test + node:assert/strict (Node built-in) | Zero dependencies, matches no-build-step constraint, fully supported on Node 25. | No |
| 010 | 2026-03-16 | auth | Testability pattern for network calls | Injectable fetchFn parameter with default to global fetch | Avoids global mock pollution, works cleanly with ESM, enables parallel test execution. | No |
| 011 | 2026-03-16 | client | SP REST hostname extraction method | `new URL(siteUrl).origin` instead of string manipulation | Handles all SharePoint subdomain variants (contoso, contoso-admin, contoso-my) correctly. Returns clean origin without path segments. Standard URL API, no regex needed. | No |
| 012 | 2026-03-16 | tools | Root site path for getSiteByUrl | Pass empty string, not `"/"` | `getSiteByUrl(hostname, "")` produces `sites/{host}:/` (correct Graph root site path). Using `"/"` would produce `sites/{host}://` (broken). | No |
| 013 | 2026-03-16 | tools | How to translate German tool descriptions | Full rewrite of tools.js instead of 50+ surgical edits | Lower risk of missed German fragments. Also covered parameter descriptions and response strings, not just tool descriptions. | No |
| 014 | 2026-03-16 | tools | Cross-module imports in tools.js | Import `discoverTenantId` from `auth.js` into `tools.js` for `connect_to_site` | Establishes pattern for tools needing auth-layer functions. First cross-module import in tools.js beyond zod. | No |
| 015 | 2026-03-16 | test | Testability of `discoverTenantId` in tool handlers | Add optional `overrides` parameter to `registerTools(server, client, auth, overrides)` | `discoverTenantId` is a module-level import in `tools.js` — can't be mocked without `--experimental-test-module-mocks`. An `overrides` bag is one line of code and avoids experimental flags. Consistent with injectable `fetchFn` pattern (D010). | No |
