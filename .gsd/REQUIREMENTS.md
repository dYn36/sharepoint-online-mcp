# Requirements

## Active

### R001 — Zero-config Auth via Well-Known Client ID
- Class: core-capability
- Status: active
- Description: Server authenticates using a pre-registered Microsoft first-party client ID (e.g. Microsoft Office `d3590ed6-52b3-4102-aeff-aad2292ab01c`). No Azure Portal, no app registration, no env vars required.
- Why it matters: The entire zero-config promise depends on this. Without it, every user needs Azure AD admin access.
- Source: user
- Primary owning slice: M001/S01
- Supporting slices: none
- Validation: unmapped
- Notes: Risk — some tenants may block device code flow via Conditional Access. Well-known client IDs may not have all required Graph scopes pre-consented.

### R002 — Auto Tenant Discovery from SharePoint URL
- Class: core-capability
- Status: active
- Description: Given a SharePoint URL like `https://contoso.sharepoint.com/sites/marketing`, the server automatically discovers the tenant ID via OpenID configuration endpoint — no manual tenant ID entry.
- Why it matters: Eliminates the need for users to know or configure their tenant ID.
- Source: user
- Primary owning slice: M001/S01
- Supporting slices: none
- Validation: unmapped
- Notes: Derive domain from URL, hit `https://login.microsoftonline.com/{domain}/.well-known/openid-configuration`, extract tenant ID from token endpoint.

### R003 — Device Code Flow Authentication
- Class: primary-user-loop
- Status: active
- Description: User authenticates by opening a URL and entering a code — no browser redirect, no localhost callback. Device code prompt appears on stderr (MCP stdio compatible).
- Why it matters: Device code flow works in headless/CLI environments where redirect-based auth doesn't.
- Source: user
- Primary owning slice: M001/S01
- Supporting slices: none
- Validation: unmapped
- Notes: Already partially implemented in existing auth.js.

### R004 — Token Caching with Silent Renewal
- Class: quality-attribute
- Status: active
- Description: Access tokens are cached locally. On subsequent requests, tokens are acquired silently from cache without re-prompting. Cache persists across server restarts.
- Why it matters: Without caching, every server restart requires re-authentication.
- Source: inferred
- Primary owning slice: M001/S01
- Supporting slices: none
- Validation: unmapped
- Notes: Already partially implemented. Cache path: `~/.sharepoint-mcp-cache.json`.

### R005 — Logout/Disconnect Tool
- Class: operability
- Status: active
- Description: An MCP tool `disconnect` clears the token cache and allows re-authentication with a different account.
- Why it matters: Users need to switch accounts or re-auth when tokens are invalid.
- Source: user
- Primary owning slice: M001/S01
- Supporting slices: none
- Validation: unmapped
- Notes: none

### R006 — Interactive Site Discovery via MCP Tools
- Class: primary-user-loop
- Status: active
- Description: Claude can search for SharePoint sites the user has access to, list followed sites, and resolve a site from its URL — all via MCP tools.
- Why it matters: Users shouldn't need to know site IDs. Discovery is the first step of every workflow.
- Source: user
- Primary owning slice: M001/S02
- Supporting slices: none
- Validation: S02 — `connect_to_site` tool registered and tested (URL parsing 7/7 tests), `list_my_sites` tool registered, `search_sites` pre-existing. Live validation deferred to milestone UAT.
- Notes: `connect_to_site` chains URL parsing → tenant discovery → auth → site resolution. `list_my_sites` calls `/me/followedSites`. `search_sites` was already present.

### R007 — Multi-Site Support
- Class: primary-user-loop
- Status: active
- Description: Users can work with multiple SharePoint sites in a single session — switching between sites without reconnecting.
- Why it matters: Many users manage multiple sites. Session-locked single-site would be limiting.
- Source: user
- Primary owning slice: M001/S02
- Supporting slices: none
- Validation: S02 — `connect_to_site` resolves any site by URL without session state; all tools accept siteId as parameter. Live multi-site validation deferred to milestone UAT.
- Notes: No "current site" concept — each tool call takes siteId explicitly. `connect_to_site` returns siteId for use in downstream calls.

### R008 — Page CRUD
- Class: core-capability
- Status: validated
- Description: Create, read, update, publish, and delete SharePoint modern pages via MCP tools.
- Why it matters: Pages are the primary content unit in SharePoint.
- Source: user
- Primary owning slice: M001/S03
- Supporting slices: none
- Validation: S03 — Contract-level: 7 mock tests prove list_pages, get_page, create_page, create_page+autoPublish, update_page, publish_page, delete_page all delegate correctly with valid MCP content shape. Live validation deferred to milestone UAT.
- Notes: Already implemented in existing tools. Auth layer wiring validated via mock delegation.

### R009 — Canvas Layout Editing
- Class: core-capability
- Status: validated
- Description: Add and configure page sections with various column layouts (full-width, 1-col, 2-col, 3-col, asymmetric).
- Why it matters: Sections are the structural building blocks of SharePoint pages.
- Source: user
- Primary owning slice: M001/S03
- Supporting slices: none
- Validation: S03 — Contract-level: 3 mock tests prove add_section (with template and emphasis) and get_page_layout delegate correctly. Live validation deferred to milestone UAT.
- Notes: Already implemented. Auth layer wiring validated via mock delegation.

### R010 — Web Part Management
- Class: core-capability
- Status: validated
- Description: Add text, image, spacer, divider, hero, quick links, and custom web parts to page sections.
- Why it matters: Web parts are how content gets placed on SharePoint pages.
- Source: user
- Primary owning slice: M001/S03
- Supporting slices: none
- Validation: S03 — Contract-level: 6 mock tests prove text, image, spacer, divider, custom webpart delegation and position shape. Live validation deferred to milestone UAT.
- Notes: Already implemented. Auth layer wiring validated via mock delegation.

### R011 — Navigation Management
- Class: core-capability
- Status: validated
- Description: Read and modify Quick Launch and Top Navigation bar entries on SharePoint sites.
- Why it matters: Navigation is a core site design element.
- Source: user
- Primary owning slice: M001/S03
- Supporting slices: none
- Validation: S03 — Contract-level: 4 mock tests prove get quick nav, get top nav, add link, remove link all call SP REST client methods (not Graph). Live dual-audience token validation deferred to milestone UAT.
- Notes: Uses SP REST API (not Graph). SP REST delegation path confirmed via mock tests.

### R012 — Branding (Logo, Asset Upload)
- Class: core-capability
- Status: validated
- Description: Set site logo and upload assets (images, files) to Site Assets library.
- Why it matters: Branding is part of the site design workflow.
- Source: user
- Primary owning slice: M001/S03
- Supporting slices: none
- Validation: S03 — Contract-level: 2 mock tests prove set_site_logo and upload_asset delegate correctly, including base64-to-Buffer decode. Live validation deferred to milestone UAT.
- Notes: Logo uses SP REST API. Upload uses Graph API.

### R013 — npm Package `sharepoint-online-mcp`
- Class: launchability
- Status: active
- Description: Published as npm package. `npx sharepoint-online-mcp` starts the server immediately. Correct bin entry, shebang, dependencies.
- Why it matters: Zero-config means zero-setup. npm is the distribution mechanism.
- Source: user
- Primary owning slice: M001/S04
- Supporting slices: none
- Validation: unmapped
- Notes: Package name: `sharepoint-online-mcp`.

### R014 — No Hardcoded Secrets
- Class: constraint
- Status: validated
- Description: No tenant IDs, client IDs, or secrets hardcoded in source or config files. All sensitive values are well-known public constants or derived at runtime.
- Why it matters: Reusability across tenants and users. Security.
- Source: user
- Primary owning slice: M001/S01
- Supporting slices: M001/S04
- Validation: S01 — grep confirms zero process.env.SHAREPOINT references in src/. Well-known client ID is a public Microsoft constant.
- Notes: Well-known client IDs are public and documented by Microsoft — they are constants, not secrets.

### R015 — Clear Error Messages
- Class: failure-visibility
- Status: active
- Description: Auth failures, permission errors, and API errors produce actionable messages explaining what went wrong and what the user can do.
- Why it matters: Zero-config means users can't debug via configuration. Error messages are the only diagnostic surface.
- Source: inferred
- Primary owning slice: M001/S04
- Supporting slices: M001/S01
- Validation: unmapped
- Notes: Especially important for Conditional Access blocks, consent failures, and scope issues.

## Deferred

### R016 — Conditional Access Fallback
- Class: continuity
- Status: deferred
- Description: If device code flow is blocked by tenant policy, offer an alternative auth method (e.g. interactive browser auth, authorization code flow with localhost redirect).
- Why it matters: Some enterprise tenants block device code flow entirely.
- Source: research
- Primary owning slice: none
- Supporting slices: none
- Validation: unmapped
- Notes: Deferred because it requires significantly more complexity (localhost server, browser launch). Revisit if device code blocking is common in practice.

## Out of Scope

### R017 — Admin Features
- Class: anti-feature
- Status: out-of-scope
- Description: Site designs, hub sites, tenant-level themes, site provisioning, admin center features.
- Why it matters: Prevents scope creep. This tool is for site content editing, not tenant administration.
- Source: user
- Primary owning slice: none
- Supporting slices: none
- Validation: n/a
- Notes: Would require admin consent and different permission scopes.

### R018 — Azure App Registration
- Class: anti-feature
- Status: out-of-scope
- Description: Requiring users to register an Azure AD application.
- Why it matters: This is the core constraint — zero Azure Portal involvement.
- Source: user
- Primary owning slice: none
- Supporting slices: none
- Validation: n/a
- Notes: The entire architecture is designed around avoiding this.

## Traceability

| ID | Class | Status | Primary owner | Supporting | Proof |
|---|---|---|---|---|---|
| R001 | core-capability | active | M001/S01 | none | unmapped |
| R002 | core-capability | active | M001/S01 | none | unmapped |
| R003 | primary-user-loop | active | M001/S01 | none | unmapped |
| R004 | quality-attribute | active | M001/S01 | none | unmapped |
| R005 | operability | active | M001/S01 | none | unmapped |
| R006 | primary-user-loop | active | M001/S02 | none | S02 — tools registered, URL parsing tested |
| R007 | primary-user-loop | active | M001/S02 | none | S02 — stateless siteId pattern, connect_to_site resolves any URL |
| R008 | core-capability | validated | M001/S03 | none | S03 — 7 mock tests, all page tools delegate correctly |
| R009 | core-capability | validated | M001/S03 | none | S03 — 3 mock tests, layout tools delegate correctly |
| R010 | core-capability | validated | M001/S03 | none | S03 — 6 mock tests, all web part types delegate correctly |
| R011 | core-capability | validated | M001/S03 | none | S03 — 4 mock tests, SP REST nav tools delegate correctly |
| R012 | core-capability | validated | M001/S03 | none | S03 — 2 mock tests, branding tools delegate correctly |
| R013 | launchability | active | M001/S04 | none | unmapped |
| R014 | constraint | validated | M001/S01 | M001/S04 | S01 grep |
| R015 | failure-visibility | validated | M001/S04 | M001/S01 | S04 — 16 tests prove actionable error messages |
| R016 | continuity | deferred | none | none | unmapped |
| R017 | anti-feature | out-of-scope | none | none | n/a |
| R018 | anti-feature | out-of-scope | none | none | n/a |

## Coverage Summary

- Active requirements: 9
- Mapped to slices: 9
- Validated: 6
- Unmapped active requirements: 0
pe | none | none | n/a |
| R018 | anti-feature | out-of-scope | none | none | n/a |

## Coverage Summary

- Active requirements: 9
- Mapped to slices: 9
- Validated: 6
- Unmapped active requirements: 0
