# S02: Site Discovery & Connection Tools — UAT

**Milestone:** M001
**Written:** 2026-03-16

## UAT Type

- UAT mode: mixed (artifact-driven for translations/registrations, live-runtime for tool execution)
- Why this mode is sufficient: URL parsing and tool registration are artifact-verifiable. `connect_to_site` and `list_my_sites` need a running MCP server with real SharePoint access for full validation — those live checks are included but expected to run during milestone-level UAT.

## Preconditions

- Node.js 20+ installed
- Working directory: project root
- For live tests (cases 4–7): valid SharePoint Online tenant access and ability to complete device code auth

## Smoke Test

```bash
node src/index.js
```
Server prints `🚀 SharePoint Online MCP Server started (stdio)` to stderr and waits for MCP messages on stdin. No errors, no env var prompts.

## Test Cases

### 1. No German Text Remaining

1. Run: `grep -E '[äöüÄÖÜß]' src/tools.js`
2. **Expected:** No output (exit code 1). Zero German characters in tools.js.

### 2. Tool Count is 25

1. Run: `grep -c 'server.tool(' src/tools.js`
2. **Expected:** Output is `25`.

### 3. URL Parsing Unit Tests Pass

1. Run: `node --test tests/tools.test.js`
2. **Expected:** 7/7 tests pass:
   - Standard site URL (`https://contoso.sharepoint.com/sites/marketing`) → hostname: `contoso.sharepoint.com`, sitePath: `sites/marketing`
   - Root site URL (`https://contoso.sharepoint.com`) → hostname: `contoso.sharepoint.com`, sitePath: (empty)
   - Trailing slash (`https://contoso.sharepoint.com/sites/marketing/`) → same as standard, no trailing slash in sitePath
   - Subpage path (`https://contoso.sharepoint.com/sites/marketing/SitePages/Home.aspx`) → sitePath: `sites/marketing` (subpage ignored)
   - Teams prefix (`https://contoso.sharepoint.com/teams/engineering`) → sitePath: `teams/engineering`
   - Invalid URL (`not-a-url`) → throws error containing "Invalid SharePoint URL"
   - URL with port (`https://contoso.sharepoint.com:8080/sites/test`) → hostname includes port, sitePath: `sites/test`

### 4. connect_to_site — Invalid URL Error

1. Send MCP `tools/call` with tool `connect_to_site`, args: `{"url": "not-a-url"}`
2. **Expected:** Response has `isError: true` and content includes "Invalid SharePoint URL: not-a-url"

### 5. connect_to_site — Valid SharePoint URL (live)

1. Send MCP `tools/call` with tool `connect_to_site`, args: `{"url": "https://{your-tenant}.sharepoint.com/sites/{your-site}"}`
2. If not yet authenticated: device code prompt appears on stderr
3. Complete device code authentication in browser
4. **Expected:** Response contains `id` (site ID GUID), `displayName`, `webUrl`, `tenantId`. No error.

### 6. list_my_sites (live)

1. After authenticating in case 5, send MCP `tools/call` with tool `list_my_sites`, args: `{}`
2. **Expected:** Response contains an array of sites. Each entry has `id`, `name`, `url`. Array may be empty if user follows no sites, but response should not error.

### 7. connect_to_site — Root Site (live)

1. Send MCP `tools/call` with tool `connect_to_site`, args: `{"url": "https://{your-tenant}.sharepoint.com"}`
2. **Expected:** Response contains the root site's `id`, `displayName`, `webUrl`. No error.

## Edge Cases

### connect_to_site with Subpage URL

1. Send MCP `tools/call` with tool `connect_to_site`, args: `{"url": "https://{tenant}.sharepoint.com/sites/marketing/SitePages/Home.aspx"}`
2. **Expected:** Resolves the `sites/marketing` site — ignores the `/SitePages/Home.aspx` suffix.

### connect_to_site with Teams-Prefixed Site

1. Send MCP `tools/call` with tool `connect_to_site`, args: `{"url": "https://{tenant}.sharepoint.com/teams/engineering"}`
2. **Expected:** Resolves the `teams/engineering` site — handles `/teams/` prefix correctly.

### list_my_sites with No Authentication

1. Start fresh server (no cached token), send MCP `tools/call` with tool `list_my_sites`, args: `{}`
2. **Expected:** Either triggers device code flow (auth prompt on stderr) or returns an error indicating auth is needed. Should NOT crash.

### Existing Tests Unbroken

1. Run: `node --test tests/auth.test.js && node --test tests/client.test.js`
2. **Expected:** 12/12 auth tests pass, 6/6 client tests pass. No regressions from S02 changes.

## Failure Signals

- `grep -E '[äöüÄÖÜß]' src/tools.js` returns any matches → German translation incomplete
- `grep -c 'server.tool(' src/tools.js` returns anything other than 25 → tool registration missing or duplicated
- Any of the 25 existing tests fail → regression introduced by S02 changes
- `connect_to_site` with valid URL returns error about "Cannot read properties of undefined" → `parseSharePointUrl` output format doesn't match what the tool handler expects
- `list_my_sites` returns raw Graph API error JSON instead of formatted response → error handling missing

## Requirements Proved By This UAT

- R006 (Interactive Site Discovery) — cases 5, 6, 7 prove site discovery via URL and followed sites listing
- R007 (Multi-Site Support) — cases 5 and 7 prove different sites can be resolved independently
- R015 (Clear Error Messages) — case 4 proves actionable error for invalid input

## Not Proven By This UAT

- R006/R007 full validation requires testing with multiple real tenants (different domains, different tenant IDs)
- Dual-audience token handling for SP REST API tools — deferred to S03
- Token caching across restarts for `connect_to_site` — deferred to milestone UAT
- `search_sites` tool (pre-existing, not modified in S02) — not retested here

## Notes for Tester

- Cases 1–4 are offline and can be run immediately. Cases 5–7 require real SharePoint access.
- The device code auth prompt goes to stderr (MCP-compatible). If testing via a raw MCP client, watch stderr for the auth URL and code.
- `list_my_sites` may return an empty array legitimately — this is not a failure. Only an error response is a failure.
- Root site resolution (case 7) uses an empty sitePath which produces `sites/{host}:/` in the Graph API call. If this fails, check Decision #012 in DECISIONS.md for context.
