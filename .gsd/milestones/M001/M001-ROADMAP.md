# M001: Zero-Config SharePoint MCP

**Vision:** `npx sharepoint-online-mcp` — no config, no Azure Portal, no env vars. Authenticate in browser, edit SharePoint sites via Claude.

## Success Criteria

- Server starts with `npx sharepoint-online-mcp` without any env vars or config files
- Device code flow authenticates against a real Azure AD tenant using well-known client ID
- Tenant ID is auto-discovered from SharePoint URL — user never enters it manually
- All 20+ existing tools work with the new auth layer (Graph API and SP REST API)
- Token persists across server restarts — no re-auth unless logout
- `npx sharepoint-online-mcp` installs and runs cleanly from npm

## Key Risks / Unknowns

- Well-known client ID scope availability — Microsoft Office client may not have pre-consented Graph Beta scopes for pages API. If not, we need to try alternative client IDs or fall back.
- Dual-audience tokens — Graph API needs `https://graph.microsoft.com` audience, SP REST needs `https://{tenant}.sharepoint.com` audience. MSAL may need to acquire two separate tokens.
- Conditional Access blocking — Some tenants block device code flow entirely. No workaround in scope.

## Proof Strategy

- Scope availability → retire in S01 by proving a token acquired with well-known client ID can call Graph Beta sites/pages endpoint
- Dual-audience → retire in S01 by proving both Graph and SP REST calls succeed (may need two token acquisitions)
- Conditional Access → documented as known limitation in S04 README

## Verification Classes

- Contract verification: Unit tests for tenant discovery URL parsing, integration test for auth flow mock
- Integration verification: Real Graph API call with acquired token, real SP REST call for navigation
- Operational verification: Server start → auth → tool call → result cycle, token cache persistence across restart
- UAT / human verification: Full workflow — start server, authenticate, search sites, create page, publish

## Milestone Definition of Done

This milestone is complete only when all are true:

- Auth module acquires tokens via well-known client ID + device code flow without any env vars
- Tenant discovery works from arbitrary SharePoint URLs
- All existing tools (pages, layout, web parts, navigation, branding) work with new auth
- Token caching works across server restarts
- Disconnect tool clears auth and enables re-auth
- `npx sharepoint-online-mcp` starts and runs cleanly
- package.json is publish-ready with correct name, bin, dependencies

## Requirement Coverage

- Covers: R001, R002, R003, R004, R005, R006, R007, R008, R009, R010, R011, R012, R013, R014, R015
- Partially covers: none
- Leaves for later: R016 (Conditional Access fallback)
- Orphan risks: none

## Slices

- [ ] **S01: Zero-Config Auth Engine** `risk:high` `depends:[]`
  > After this: Server starts without env vars, device code authenticates against a real tenant using well-known client ID, token is cached, Graph API call succeeds — proven against live Azure AD
- [ ] **S02: Site Discovery & Connection Tools** `risk:medium` `depends:[S01]`
  > After this: Claude can search sites, resolve a site from URL with auto tenant discovery, list followed sites, and disconnect/reconnect — all via MCP tools against live SharePoint
- [ ] **S03: Tool Validation & Dual-Audience Tokens** `risk:medium` `depends:[S02]`
  > After this: All 20+ existing tools (pages, layout, web parts, navigation, branding) work with the new auth layer — both Graph API and SP REST API calls succeed with correct token audiences
- [ ] **S04: npm Package & Polish** `risk:low` `depends:[S03]`
  > After this: `npx sharepoint-online-mcp` installs and runs cleanly, README documents zero-config workflow, error messages are actionable, package is publish-ready

## Boundary Map

### S01 → S02

Produces:
- `SharePointAuth` class with `getAccessToken(resource)` that handles well-known client ID, device code flow, token caching, and silent renewal
- `discoverTenantId(sharepointDomain)` function that resolves tenant ID from SharePoint URL
- Server starts without env vars, auth is triggered lazily on first token request

Consumes:
- nothing (first slice)

### S02 → S03

Produces:
- `connect_to_site` tool that takes a SharePoint URL, auto-discovers tenant, triggers auth if needed, resolves site ID
- `search_sites`, `list_my_sites`, `disconnect` MCP tools
- Site context is available for downstream tools

Consumes:
- Auth engine from S01 (token acquisition, tenant discovery)

### S03 → S04

Produces:
- All existing tools validated and working with new auth
- `SharePointClient` updated for dual-audience token support (Graph vs SP REST)
- Any tool fixes needed for well-known client ID tokens

Consumes:
- Auth engine from S01, site connection tools from S02

### S04 (terminal)

Produces:
- `package.json` with name `sharepoint-online-mcp`, correct bin/dependencies
- README with zero-config workflow documentation
- Actionable error messages for common failure modes

Consumes:
- Everything from S01-S03
