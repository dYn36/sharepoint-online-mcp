# M001: Zero-Config SharePoint MCP

**Gathered:** 2026-03-16
**Status:** Ready for planning

## Project Description

Refactor an existing SharePoint MCP server (20+ tools) from requiring Azure AD app registration + env vars to a completely zero-config experience using well-known Microsoft first-party client IDs and automatic tenant discovery.

## Why This Milestone

The current server requires users to register an Azure AD app, configure permissions, and set env vars — a process that most users can't do without admin help. The refactored server should work immediately with `npx sharepoint-online-mcp` and no configuration at all.

## User-Visible Outcome

### When this milestone is complete, the user can:

- Run `npx sharepoint-online-mcp` and have the server start immediately — no env vars, no config
- On first SharePoint tool call, see a device code prompt, authenticate in browser, and be connected
- Ask Claude to find SharePoint sites, pick one, and start editing pages/navigation/branding
- Switch between multiple SharePoint sites in one session
- Disconnect and re-authenticate with a different account

### Entry point / environment

- Entry point: `npx sharepoint-online-mcp` or MCP client config (Claude Desktop, etc.)
- Environment: local dev, CLI, MCP stdio
- Live dependencies involved: Microsoft Graph API, SharePoint REST API, Azure AD OAuth

## Completion Class

- Contract complete means: Server starts without env vars, auth module resolves tenant and acquires token, all tools compile and register
- Integration complete means: Device code flow works against a real Azure AD tenant, Graph API calls succeed with the acquired token, SP REST API calls succeed for navigation/branding
- Operational complete means: Token caching works across restarts, silent renewal works, logout clears cache and re-auth works

## Final Integrated Acceptance

To call this milestone complete, we must prove:

- `npx sharepoint-online-mcp` starts cleanly, Claude calls a tool, device code appears, user authenticates, and a real SharePoint page is created/edited/published
- Navigation and branding tools work (they use SP REST API, not Graph — different token audience)
- Token cache persists across server restart (no re-auth on second launch)
- Disconnect tool clears auth and enables re-authentication

## Risks and Unknowns

- Well-known client ID may not have sufficient pre-consented scopes for Graph Beta pages API — this is the highest risk
- SP REST API may require a different token audience than Graph API — need dual-audience token strategy
- Tenant Conditional Access may block Device Code Flow — cannot be worked around without alternative auth
- Graph Beta pages API may behave differently with well-known client tokens vs custom app tokens

## Existing Codebase / Prior Art

- `src/auth.js` — MSAL Device Code Flow, token cache, 95 lines. Needs: remove client ID param, add tenant discovery, add well-known client ID
- `src/client.js` — Graph + SP REST client, 225 lines. Needs: dual-token support (Graph audience + SP audience), lazy auth trigger
- `src/tools.js` — 20+ MCP tools, 530 lines. Needs: add disconnect tool, add connect/site-context tools. Existing tools are fine.
- `src/index.js` — Server entrypoint, 55 lines. Needs: remove env var requirements, lazy init
- `src/auth-cli.js` — Standalone auth test, 30 lines. Needs: update for zero-config

> See `.gsd/DECISIONS.md` for all architectural and pattern decisions.

## Relevant Requirements

- R001-R005 — Auth foundation (zero-config, tenant discovery, device code, caching, logout)
- R006-R007 — Site discovery and multi-site support
- R008-R012 — Page/layout/webpart/navigation/branding tools (validation)
- R013-R015 — npm package, no hardcoding, error messages

## Scope

### In Scope

- Auth refactor to well-known client ID + auto tenant discovery
- Lazy authentication on first tool use
- Disconnect/logout MCP tool
- Site search and connection tools
- Validation of all existing tools with new auth
- npm package `sharepoint-online-mcp` with bin entry
- Clear error messages for auth/permission failures

### Out of Scope / Non-Goals

- Azure AD app registration flow
- Admin-level features (site provisioning, hub sites, tenant themes)
- TypeScript migration
- Alternative auth flows (browser redirect, auth code)
- SharePoint on-premises support

## Technical Constraints

- Must use well-known first-party client ID — no custom app registration
- Must work over MCP stdio (no interactive terminal for auth — device code on stderr)
- No build step — plain ESM JavaScript, npx-ready
- Node.js >= 18 (fetch available globally in newer versions, but keeping node-fetch for compatibility)

## Integration Points

- Microsoft Graph API v1.0 + Beta — pages, sites, lists, files
- SharePoint REST API — navigation, branding, theming
- Azure AD OAuth 2.0 — device code flow, token cache
- MSAL Node — token acquisition, cache management
- MCP SDK — tool registration, stdio transport

## Open Questions

- Will `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Office client) get `Sites.ReadWrite.All` scope? Need to test. Fallback: try Azure CLI client ID.
- Does SP REST API accept tokens acquired with Graph resource scope, or do we need separate token acquisition with SP resource scope? Need to test.
