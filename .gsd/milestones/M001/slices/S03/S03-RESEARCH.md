# S03: Tool Validation & Dual-Audience Tokens — Research

**Date:** 2026-03-16

## Summary

The dual-audience token routing is already implemented and tested. `client.js` passes `"https://graph.microsoft.com"` for Graph calls and `new URL(siteUrl).origin` for SP REST calls. `auth.getAccessToken(resource)` builds `{resource}/.default` scopes. Six unit tests in `client.test.js` verify this routing. The 25 registered tools all call through `client.js` correctly.

The real work in S03 is **tool handler test coverage** and **verifying the end-to-end wiring** — ensuring that every tool category (pages, layout, web parts, navigation, branding, files) correctly delegates to the client and that the client constructs valid API calls. Currently only `parseSharePointUrl` has tool-level tests; the 25 tool handler callbacks have zero test coverage.

A secondary concern is whether MSAL's `acquireTokenSilent` can exchange a refresh token from Graph-audience auth for an SP-audience token without triggering a second device code prompt. This is an MSAL capability question — the code is wired correctly, but it can only be validated against a live tenant (out of scope for unit tests, in scope for milestone UAT).

## Recommendation

Build a comprehensive tool handler test suite using a mock server + mock client pattern. This validates:
1. Each tool is registered and callable
2. Each tool passes correct arguments to the client
3. Graph-based tools route through `graph()`/`graphBetaReq()` (Graph audience)
4. SP REST tools route through `spRest()` (SP audience)
5. Error handling returns `isError: true` with messages
6. Response format is correct MCP content

No changes to `client.js` or `auth.js` are needed — the dual-audience wiring is complete and tested. `tools.js` itself needs no changes — it's already correct. The deliverable is **validation through tests**, not code changes.

## Implementation Landscape

### Key Files

- `src/tools.js` (530 lines) — 25 tool registrations. All tool handlers live here. No changes needed, only test coverage.
- `src/client.js` (225 lines) — `request()` for Graph, `spRest()` for SP REST. Already tested for token routing. No changes needed.
- `src/auth.js` (135 lines) — `getAccessToken(resource)` with `buildScopes()`. Already tested. No changes needed.
- `tests/tools.test.js` — Currently 7 tests (URL parsing only). Needs expansion to cover tool handlers.
- `tests/client.test.js` — 6 tests for token routing. Complete for S03's needs.

### Tool → API Mapping

**Graph API tools** (audience: `https://graph.microsoft.com`):
| Tool | Client Method | Graph Endpoint |
|------|--------------|----------------|
| `search_sites` | `listSites(query)` | `GET /sites?search=` |
| `get_site_details` | `getSite()` + `getSiteLists()` | `GET /sites/{id}`, `GET /sites/{id}/lists` |
| `get_site_by_url` | `getSiteByUrl()` | `GET /sites/{host}:/{path}` |
| `connect_to_site` | `getSiteByUrl()` + `discoverTenantId()` | `GET /sites/{host}:/{path}` |
| `list_my_sites` | `listSites("")` | `GET /me/followedSites` |
| `list_pages` | `listPages()` | `GET beta /sites/{id}/pages` |
| `get_page` | `getPage()` | `GET beta /sites/{id}/pages/{pid}` |
| `create_page` | `createPage()` | `POST beta /sites/{id}/pages` |
| `update_page` | `updatePage()` | `PATCH beta /sites/{id}/pages/{pid}` |
| `publish_page` | `publishPage()` | `POST beta .../publish` |
| `delete_page` | `deletePage()` | `DELETE beta /sites/{id}/pages/{pid}` |
| `add_section` | `addHorizontalSection()` | `POST beta .../horizontalSections` |
| `get_page_layout` | `getHorizontalSections()` + `getPageWebParts()` | `GET beta .../horizontalSections`, `GET beta .../webParts` |
| `add_text_webpart` | `createWebPartInSection()` | `POST beta .../webparts` |
| `add_image_webpart` | `createWebPartInSection()` | `POST beta .../webparts` |
| `add_spacer` | `createWebPartInSection()` | `POST beta .../webparts` |
| `add_divider` | `createWebPartInSection()` | `POST beta .../webparts` |
| `add_custom_webpart` | `createWebPartInSection()` | `POST beta .../webparts` |
| `upload_asset` | `uploadFile()` | `PUT /sites/{id}/drive/root:/.../content` |
| `get_design_templates` | (none — static data) | (none) |

**SP REST tools** (audience: `https://{tenant}.sharepoint.com`):
| Tool | Client Method | SP REST Endpoint |
|------|--------------|-----------------|
| `get_navigation` | `getNavigation()` / `getTopNavigation()` | `_api/web/navigation/quicklaunch` or `topnavigationbar` |
| `add_navigation_link` | `addNavigationNode()` | `POST _api/web/navigation/...` |
| `remove_navigation_link` | `deleteNavigationNode()` | `DELETE _api/web/navigation/...(id)` |
| `set_site_logo` | `setSiteLogo()` | `POST _api/web` (MERGE) |

**Auth tools** (no API call):
| Tool | Action |
|------|--------|
| `disconnect` | `auth.logout()` — clears cache |

### Build Order

1. **Create mock server + mock client test infrastructure** — a lightweight mock of `McpServer` that records `tool()` registrations and allows calling handlers, plus a mock client that records method calls and returns canned responses. This is the foundation for all tool tests.

2. **Test Graph-based tool handlers** (pages, layout, web parts, sites, files) — verify each handler calls the right client method with correct args, returns proper MCP content format. This covers R008, R009, R010, R012.

3. **Test SP REST tool handlers** (navigation, branding) — verify these handlers pass `siteUrl` to client methods that route through `spRest()`. This covers R011, R012 and validates the dual-audience path from tool → client → auth.

4. **Test disconnect tool** — verify it calls `auth.logout()`. Covers R005.

5. **Test error handling** — verify `connect_to_site` and `list_my_sites` return `isError: true` on failures.

### Verification Approach

- `node --test tests/*.test.js` — all tests pass (existing 25 + new tool handler tests)
- Each tool category has at least one representative test: sites, pages, layout, web parts, navigation, branding, files, auth
- Test confirms both Graph and SP REST paths are exercised
- `grep -c 'server.tool(' src/tools.js` still equals 25 (no regressions)
- `node src/index.js` starts without error

### Testing Pattern

The mock server pattern for testing `registerTools()`:

```javascript
// Mock server that records tool registrations
class MockServer {
  constructor() { this.tools = new Map(); }
  tool(name, desc, schema, handler) { this.tools.set(name, { desc, schema, handler }); }
  async call(name, args) { return this.tools.get(name).handler(args); }
}

// Mock client that records calls and returns canned data
class MockClient {
  constructor() { this.calls = []; }
  async listPages(siteId) { this.calls.push(['listPages', siteId]); return [...]; }
  // ... etc for each method
}
```

This avoids module mocking complexity and works cleanly with ESM + `node:test`.

## Constraints

- `client.js` uses `import fetch from "node-fetch"` at module level — cannot be replaced in tests without module mocking. The existing `client.test.js` works around this by catching fetch errors and only verifying auth calls. Tool handler tests should mock at the client level instead (pass a mock client to `registerTools`).
- `tools.js` imports `discoverTenantId` from `auth.js` — `connect_to_site` handler calls this directly, not through the client. Tool handler tests for `connect_to_site` need a way to mock `discoverTenantId`. Options: (a) accept that `connect_to_site` tests will need to mock at the module level or (b) test `connect_to_site` by providing a mock `discoverTenantId` via dependency injection refactor. Simplest approach: make `registerTools` accept an optional 4th arg for overrides, or test `connect_to_site` at a higher level.

## Common Pitfalls

- **Mocking `discoverTenantId` in `connect_to_site` tests** — It's imported at module level in `tools.js`. `node:test`'s `mock.module()` can handle this but requires `--experimental-test-module-mocks` flag. Alternative: refactor to inject it, or just test `connect_to_site`'s URL parsing and error paths without mocking tenant discovery.
- **Assuming single device code prompt** — The user may see two device code prompts if MSAL can't use the refresh token across resources. This is expected MSAL behavior with `common` authority and well-known client IDs. The code handles it correctly (silent try → device code fallback), but documentation should mention it. Not a code change — a docs note for S04.

## Open Risks

- **MSAL cross-resource silent token acquisition** — Whether `acquireTokenSilent` succeeds for SP resource after initial Graph device code auth depends on the well-known client ID having sufficient consent for both resources. Can only be validated against a live tenant. If it fails, users see a second device code prompt — the code handles this gracefully, but UX is degraded. This is a milestone UAT concern, not an S03 blocker.
