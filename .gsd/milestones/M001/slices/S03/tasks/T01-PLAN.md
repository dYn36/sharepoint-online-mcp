# T01: Tool handler test suite with mock server/client/auth

## Context

**Slice goal:** Validate all 25 MCP tool handlers work correctly with mock infrastructure — proving delegation, argument passing, response format, and error handling.

**Slice demo:** `node --test tests/*.test.js` passes 50+ tests total.

**Slice verification:**
- `node --test tests/tools.test.js` — 30+ tool handler tests pass
- `node --test tests/*.test.js` — all tests across all files pass
- `grep -c 'server.tool(' src/tools.js` still equals 25
- `node src/index.js` starts without error

## Description

Expand `tests/tools.test.js` to add comprehensive handler tests for all 25 tools. The file currently has 7 tests covering only `parseSharePointUrl`. The new tests use a mock infrastructure that intercepts `registerTools()` — since it takes `(server, client, auth)` as parameters, we can inject mocks directly.

One small source change: add an optional 4th parameter `overrides = {}` to `registerTools` so that `discoverTenantId` can be injected for testing `connect_to_site`. This is the only source code change in the entire slice.

## Inputs

- `src/tools.js` — 25 tools registered via `registerTools(server, client, auth)`. Read this file to understand each handler's arguments and return values.
- `tests/tools.test.js` — existing 7 tests for `parseSharePointUrl`. Expand this file.

## Steps

### 1. Add `overrides` parameter to `registerTools`

In `src/tools.js`, change the function signature:

**Current (line 138):**
```js
export function registerTools(server, client, auth) {
```

**New:**
```js
export function registerTools(server, client, auth, overrides = {}) {
```

Then on line 237 where `discoverTenantId` is called inside `connect_to_site`:
```js
const tenantId = await discoverTenantId(hostname);
```
Change to:
```js
const _discoverTenantId = overrides.discoverTenantId || discoverTenantId;
const tenantId = await _discoverTenantId(hostname);
```

This is the only source change. Everything else is test code.

### 2. Build mock infrastructure in tests/tools.test.js

Add these classes after the existing imports:

**MockServer** — records `tool()` registrations and allows calling handlers:
```js
class MockServer {
  constructor() { this.tools = new Map(); }
  tool(name, _desc, _schema, handler) {
    this.tools.set(name, handler);
  }
  async call(name, args) {
    const handler = this.tools.get(name);
    if (!handler) throw new Error(`Tool not registered: ${name}`);
    return handler(args);
  }
}
```

**MockClient** — stubs for all 20 client methods called by tools. Each stub records the call in `this.calls` and returns minimal canned data. The canned data must be just enough for the handler to not throw — usually an object with `id`, `displayName`, `webUrl`, etc.

Client methods to stub (return signatures from reading tool handlers):
- `listSites(query)` → `[{ id: "s1", displayName: "Site", webUrl: "https://x", description: "d" }]`
- `getSite(siteId)` → `{ id: siteId, displayName: "Site", webUrl: "https://x", description: "d", createdDateTime: "2024-01-01" }`
- `getSiteLists(siteId)` → `[{ id: "l1", displayName: "Documents", list: { template: "documentLibrary" } }]`
- `getSiteByUrl(hostname, path)` → `{ id: "s1", displayName: "Site", webUrl: "https://x", description: "d" }`
- `listPages(siteId)` → `[{ id: "p1", name: "Home.aspx", title: "Home", webUrl: "https://x/p", pageLayout: "article", lastModifiedDateTime: "2024-01-01" }]`
- `getPage(siteId, pageId)` → `{ id: pageId, title: "Page", webUrl: "https://x" }`
- `createPage(siteId, opts)` → `{ id: "newpage", name: opts.name, title: opts.title, webUrl: "https://x/p" }`
- `updatePage(siteId, pageId, updates)` → `undefined`
- `publishPage(siteId, pageId)` → `undefined`
- `deletePage(siteId, pageId)` → `undefined`
- `addHorizontalSection(siteId, pageId, template)` → `{ id: "sec1" }`
- `getHorizontalSections(siteId, pageId)` → `[{ id: "sec1", layout: "fullWidth" }]`
- `getPageWebParts(siteId, pageId)` → `[{ id: "wp1", type: "text" }]`
- `createWebPartInSection(siteId, pageId, pos, webPart)` → `{ id: "wp-new" }`
- `getNavigation(siteUrl)` → `[{ Id: 1, Title: "Home", Url: "/" }]`
- `getTopNavigation(siteUrl)` → `[{ Id: 2, Title: "About", Url: "/about" }]`
- `addNavigationNode(siteUrl, navType, node)` → `undefined`
- `deleteNavigationNode(siteUrl, navType, nodeId)` → `undefined`
- `setSiteLogo(siteUrl, logoUrl)` → `undefined`
- `uploadFile(siteId, folder, fileName, buffer)` → `{ id: "f1", webUrl: "https://x/f" }`

Each method should record `[methodName, ...args]` into `this.calls` array before returning.

**MockAuth** — `{ logout: async () => { this.logoutCalled = true; } }`

**MockDiscoverTenantId** — `async (hostname) => "fake-tenant-id-123"`

### 3. Set up test structure and register tools once per suite

Before each `describe` group (or once at the top), create instances and call `registerTools`:
```js
const server = new MockServer();
const client = new MockClient();
const auth = new MockAuth();
registerTools(server, client, auth, { discoverTenantId: mockDiscoverTenantId });
```

Between tests, clear `client.calls` to isolate assertions.

### 4. Write tests for each tool category

**Every test must verify:**
1. The handler can be called (it's registered)
2. It calls the correct client method(s) with expected args
3. The return value has the shape `{ content: [{ type: "text", text: "..." }] }`
4. Where applicable, `isError: true` is present on error paths

#### Sites (5 tests)
- `search_sites` — calls `listSites(query)`, response contains mapped site array
- `get_site_details` — calls `getSite(siteId)` AND `getSiteLists(siteId)` (parallel), response has site + lists
- `get_site_by_url` — calls `getSiteByUrl(hostname, sitePath)`
- `connect_to_site` — calls mock `discoverTenantId` then `getSiteByUrl`, response has tenantId
- `list_my_sites` — calls `listSites("")`, response has mapped sites

#### Pages (5 tests)
- `list_pages` — calls `listPages(siteId)`, response has page array
- `get_page` — calls `getPage(siteId, pageId)`
- `create_page` — calls `createPage(siteId, opts)`, verify name/title passed through
- `create_page` with `autoPublish: true` — calls `createPage` then `publishPage`
- `delete_page` — calls `deletePage(siteId, pageId)`

#### Layout (3 tests)
- `add_section` — calls `addHorizontalSection` with template properties
- `add_section` with emphasis — verify emphasis is included in template
- `get_page_layout` — calls both `getHorizontalSections` AND `getPageWebParts` (parallel)

#### Web Parts (6 tests)
- `add_text_webpart` — calls `createWebPartInSection` with text web part template
- `add_image_webpart` — calls `createWebPartInSection` with image web part template containing imageUrl
- `add_spacer` — calls `createWebPartInSection` with spacer template
- `add_divider` — calls `createWebPartInSection` with divider template
- `add_custom_webpart` — calls `createWebPartInSection` with parsed JSON
- Check that web part position object has `{ sectionIndex, columnIndex }` shape

#### Navigation — SP REST (3 tests)
- `get_navigation` with `navType: "quick"` — calls `getNavigation(siteUrl)` (not `getTopNavigation`)
- `get_navigation` with `navType: "top"` — calls `getTopNavigation(siteUrl)`
- `add_navigation_link` — calls `addNavigationNode(siteUrl, navType, node)`
- `remove_navigation_link` — calls `deleteNavigationNode(siteUrl, navType, nodeId)`

#### Branding (2 tests)
- `set_site_logo` — calls `setSiteLogo(siteUrl, logoUrl)` (SP REST path)
- `upload_asset` — calls `uploadFile(siteId, folder, fileName, buffer)`, verify buffer is from base64

#### Auth & utility (3 tests)
- `disconnect` — calls `auth.logout()`
- `get_design_templates` — returns static data (no client calls), response has sectionLayouts
- `connect_to_site` error path — feed it an invalid URL, verify `isError: true`
- `list_my_sites` error path — make `listSites` throw, verify `isError: true`

That's ~30 new tests. Combined with 7 existing URL parsing tests → 37+ in tools.test.js.

### 5. Run full verification

```bash
node --test tests/tools.test.js
node --test tests/auth.test.js
node --test tests/client.test.js
grep -c 'server.tool(' src/tools.js  # must be 25
node src/index.js  # must start without error (Ctrl+C after)
```

## Must-Haves

- MockServer, MockClient, MockAuth classes that enable calling any registered handler
- `overrides` parameter added to `registerTools` for `discoverTenantId` injection
- At least one test per tool category: sites, pages, layout, web parts, navigation, branding, files, auth
- Error path tests for `connect_to_site` and `list_my_sites`
- All existing 7 URL parsing tests continue to pass
- `node --test tests/*.test.js` — zero failures across all test files

## Verification

```bash
node --test tests/tools.test.js  # 30+ tests pass
node --test tests/*.test.js      # all pass (auth + client + tools)
grep -c 'server.tool(' src/tools.js  # 25
node src/index.js                # starts without error
```

## Done When

- Every tool category has ≥1 test: sites (5), pages (5), layout (3), web parts (6), navigation (4), branding (2), auth/utility (4) = ~29 new tests minimum
- All tests pass
- Source change is limited to the `overrides` parameter in `registerTools`
- No regressions in existing tests

## Expected Output

- Modified: `src/tools.js` (2-line change: overrides parameter + discoverTenantId override)
- Modified: `tests/tools.test.js` (expanded from ~39 lines to ~400+ lines with mock infra and 30+ tests)

## Observability Impact

- **New diagnostic surface:** `MockClient.calls` array pattern — any future test can inspect exact args passed to any client method.
- **Test run signal:** `node --test tests/tools.test.js` — TAP output with per-handler pass/fail. A future agent can run this single command to verify all 25 tool handlers.
- **Error shape visibility:** Error path tests assert `{ isError: true, content: [{ type: "text", text: "Error..." }] }` — this shape is now a tested contract, not just convention.
- **Injection seam:** `overrides` parameter on `registerTools` — allows injecting any dependency (currently `discoverTenantId`) without modifying source internals. Visible in function signature.
