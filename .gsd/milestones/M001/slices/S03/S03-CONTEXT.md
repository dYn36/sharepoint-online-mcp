---
id: S03
milestone: M001
status: ready
---

# S03: Tool Validation & Dual-Audience Tokens — Context

<!-- Slice-scoped context. Milestone-only sections (acceptance criteria, completion class,
     milestone sequence) do not belong here — those live in the milestone context. -->

## Goal

All 20+ existing tools work with the new zero-config auth layer — Graph API and SP REST API calls succeed with correct token audiences, tools use site context, descriptions are English with prerequisites, and a get_status tool gives Claude visibility into auth and connection state.

## Why this Slice

S01 provides auth and S02 provides site discovery, but the existing 22 tools haven't been validated against the new auth layer. Graph tools and SP REST tools need different token audiences, tool descriptions are in German, and none of them use the site context from S02. This slice makes the full tool set usable end-to-end. S04 (npm packaging) depends on a validated, working tool set.

## Scope

### In Scope

- Dual-audience token support transparent in `client.js` — `graph()`/`graphBetaReq()` automatically use Graph audience, `spRest()` automatically uses SP audience. Tools don't specify audience.
- All 22 tools get site context fallback — `siteId` and `siteUrl` auto-filled from connected site; explicit parameter still accepted as override
- Translate all tool descriptions and parameter hints from German to English — concise (1 sentence) with usage hint and prerequisites (e.g. "Requires connected site")
- Organize tools into logical categories in code (Site, Pages, Layout, WebParts, Navigation, Branding)
- Add `get_status` MCP tool — returns auth state (connected/disconnected), connected site info, and available tool categories
- All tools always registered with MCP server — disabled/unavailable tools return clear error explaining why (not hidden from tool list)
- Live validation of each tool against real SharePoint tenant — manual test script, pass/fail tracked per tool
- Fix minor tool issues as needed (response format changes, parameter adjustments) — major rewrites out of scope
- If Graph Beta pages endpoints return 403/401 with well-known client ID, disable page tools gracefully with clear error — no v1.0 fallback
- Silent token renewal via MSAL — if silent fails, re-trigger device code flow transparently
- Retry once with exponential backoff on 429 (rate limit) — if still throttled, return error with explanation
- Clear, actionable error messages on scope/permission failures — explain which permission is missing and suggest next steps

### Out of Scope

- Adding new tools beyond `get_status`
- Major tool rewrites or new functionality
- Automated CI tests (requires real credentials — manual live testing only)
- v1.0 fallback for Graph Beta endpoints
- npm packaging, README (S04)
- New auth flows or client ID changes

## Constraints

- Must consume `SharePointAuth.getAccessToken(resource)` from S01 — dual-audience is handled by passing different resource strings, not by changing auth
- Must consume site context (default `siteId`/`siteUrl`) from S02 — tools fall back to connected site when parameter not provided
- Plain JavaScript ESM, no build step (D006)
- All user-facing strings in English
- Token scope strategy per D007: `.default` for both Graph and SP REST audiences

## Integration Points

### Consumes

- `SharePointAuth.getAccessToken(resource)` — S01 auth engine, called with `https://graph.microsoft.com/.default` or `https://{tenant}.sharepoint.com/.default`
- `discoverTenantId(domain)` — S01 tenant discovery, needed to construct SP REST audience URL
- Site context from S02 — default `siteId`, `siteUrl`, `tenantDomain` for auto-filling tool parameters
- `connect_to_site`, `disconnect` from S02 — site lifecycle tools
- Existing `src/client.js` — refactored for dual-audience token acquisition
- Existing `src/tools.js` — 22 tools refactored for site context, English descriptions, error handling

### Produces

- `SharePointClient` updated with transparent dual-audience token support (`graph()` → Graph token, `spRest()` → SP token)
- All 22 existing tools validated and working (or gracefully disabled) with new auth
- `get_status` MCP tool — auth state, connected site, available tool categories
- English tool descriptions with prerequisites and usage hints
- Tools organized into categories (Site, Pages, Layout, WebParts, Navigation, Branding)
- Manual live test script for tool validation
- Retry-once-with-backoff for 429 rate limits in client layer

## Open Questions

- Which tools will actually work with the well-known client ID? — Current thinking: Graph v1.0 site/list tools very likely work; Graph Beta pages tools are highest risk; SP REST navigation/branding tools depend on SP audience token working. Live testing will determine.
- Does retry-once on 429 need a configurable delay, or is a fixed 2-5 second backoff sufficient? — Current thinking: fixed backoff is fine for interactive use.
- Should `get_status` also report which specific tools are disabled and why? — Current thinking: yes, include a list of disabled tools with reasons so Claude can tell the user.
