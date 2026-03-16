---
id: S02
milestone: M001
status: ready
---

# S02: Site Discovery & Connection Tools — Context

<!-- Slice-scoped context. Milestone-only sections (acceptance criteria, completion class,
     milestone sequence) do not belong here — those live in the milestone context. -->

## Goal

Claude can search sites, connect to a site from any SharePoint URL with auto tenant discovery and auto auth, list followed sites, and disconnect — all via MCP tools against live SharePoint, with a persistent default site context so downstream tools don't need siteId on every call.

## Why this Slice

S01 delivers the auth engine but no way to discover or select a site. Every existing tool requires an explicit `siteId` parameter, which Claude has to obtain somehow. This slice provides the site discovery and connection workflow that makes the existing 20+ tools usable: search → pick → connect → work. It also introduces the site context model that S03 depends on for tool validation.

## Scope

### In Scope

- `connect_to_site` MCP tool — accepts any SharePoint URL (page, document, list, site root), extracts site path, auto-discovers tenant, triggers auth if needed, resolves site ID via Graph, sets as default site context
- `search_sites` MCP tool — refactor existing tool to work with new auth (already exists in `src/tools.js`, needs auth layer update)
- `list_my_sites` MCP tool — list sites the user follows (existing `followedSites` Graph call in `src/client.js`)
- `disconnect` MCP tool — clears site context AND current tenant's auth token (per S01 context: uses MSAL `removeAccount()` for active account, not full cache wipe)
- Default site context — connecting to a site sets a persistent default so existing tools don't require `siteId` on every call; tools still accept explicit `siteId` to override
- One site at a time — connecting to a new site silently replaces the previous context, no confirmation needed
- Cross-tenant switching — if new site is on a different tenant, re-auth triggers automatically
- Flexible URL parsing — extract site from any SharePoint URL format (site root, page URL, document library URL, list URL, etc.)

### Out of Scope

- Modifying existing tool implementations beyond adding default site context fallback (S03)
- Dual-audience token handling for SP REST tools (S03)
- npm packaging, README (S04)
- Site provisioning, hub site management, or admin features
- Persisting site context across server restarts (in-memory only for now)
- Multi-site simultaneous connections

## Constraints

- Must consume `SharePointAuth` with `getAccessToken(resource)` and `discoverTenantId()` from S01 — no reimplementation
- Plain JavaScript ESM, no build step (D006)
- MCP stdio transport — all user feedback via tool results, not interactive prompts
- `connect_to_site` is the primary entry point for the zero-config experience — it must feel like one step: paste URL → authenticated → connected
- Disconnect clears both site context and auth for current tenant (not all tenants)

## Integration Points

### Consumes

- `SharePointAuth.getAccessToken(resource)` — token acquisition from S01
- `SharePointAuth.discoverTenantId(domain)` — tenant resolution from S01
- `SharePointAuth` account management — `removeAccount()` for disconnect
- `SharePointClient.getSiteByUrl(hostname, sitePath)` — existing Graph call to resolve site from URL
- `SharePointClient.listSites(query)` — existing Graph search
- `SharePointClient.graph("/me/followedSites")` — existing followed sites call

### Produces

- `connect_to_site` MCP tool — one-step site connection from any SharePoint URL
- `search_sites` MCP tool — search sites by keyword (refactored from existing)
- `list_my_sites` MCP tool — list followed/accessible sites
- `disconnect` MCP tool — clear site context + current tenant auth
- Site context state — default `siteId` and `siteUrl` available for downstream tools
- URL parsing utility — extracts hostname and site path from arbitrary SharePoint URLs

## Open Questions

- How deep should URL parsing go? — Current thinking: handle `/sites/xxx` and `/teams/xxx` paths, plus root site. Deeply nested URLs (e.g. page URLs with `/SitePages/Home.aspx`) get truncated to the site root. Edge cases discovered during implementation get added.
- Should `list_my_sites` require auth first, or also auto-trigger? — Current thinking: auto-trigger, same as `connect_to_site`. Any tool that needs a token triggers auth lazily.
- What happens if `connect_to_site` URL points to a site the user doesn't have access to? — Current thinking: let the Graph API error propagate as a clear error message (e.g. "Access denied to this site. You may need to request access from the site owner.").
