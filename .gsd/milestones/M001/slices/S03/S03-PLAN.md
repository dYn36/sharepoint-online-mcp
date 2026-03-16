# S03: Tool Validation & Dual-Audience Tokens

**Goal:** All 25 existing tools are validated to correctly delegate to the client layer with proper arguments, and the dual-audience token routing (Graph vs SP REST) is proven through test coverage.
**Demo:** `node --test tests/*.test.js` passes 50+ tests covering every tool category — pages, layout, web parts, navigation, branding, files, auth — with mock infrastructure proving each tool calls the right client method with correct arguments and returns valid MCP content.

## Must-Haves

- Every tool handler in `registerTools()` has at least one test proving correct client method delegation
- Graph-based tools (pages, layout, web parts, sites, files) call client methods that route through Graph audience
- SP REST tools (navigation, branding) call client methods that route through SP REST audience
- `disconnect` tool calls `auth.logout()`
- Error paths return `isError: true` with actionable messages
- All response formats are valid MCP content (`{ content: [{ type: "text", text: ... }] }`)
- Existing 25 tests (URL parsing + auth + client) continue to pass — zero regressions

## Proof Level

- This slice proves: contract
- Real runtime required: no — mock-based unit tests validate wiring
- Human/UAT required: no — live dual-audience token behavior deferred to milestone UAT

## Verification

- `node --test tests/tools.test.js` — 30+ tool handler tests pass (all categories represented)
- `node --test tests/*.test.js` — all tests across all files pass (auth, client, tools)
- `grep -c 'server.tool(' src/tools.js` still equals 25 (no source regressions)
- `node src/index.js` starts without error (no import breakage)

## Integration Closure

- Upstream surfaces consumed: `registerTools(server, client, auth)` from `src/tools.js`, `discoverTenantId` from `src/auth.js`
- New wiring introduced in this slice: none — tests only, no source changes
- What remains before the milestone is truly usable end-to-end: S04 (npm packaging, README, polish)

## Tasks

- [ ] **T01: Tool handler test suite with mock server/client/auth** `est:45m`
  - Why: The 25 tool handlers have zero test coverage. This is the entire deliverable of S03 — proving that every tool correctly delegates to the client and returns valid MCP content. Covers R008, R009, R010, R011, R012.
  - Files: `tests/tools.test.js`
  - Do: Expand `tests/tools.test.js` with a mock infrastructure (MockServer that records `tool()` calls and allows invoking handlers, MockClient with stub methods that record calls and return canned data, mock auth with `logout()`) and test every tool category. See task plan for full specification.
  - Verify: `node --test tests/*.test.js` passes all tests; `grep -c 'server.tool(' src/tools.js` = 25; `node src/index.js` starts clean
  - Done when: Every tool category (sites, pages, layout, web parts, navigation, branding, files, disconnect) has at least one passing test, error paths tested, total test count ≥ 30 new tests in tools.test.js

## Observability / Diagnostics

- **Test run signal:** `node --test tests/tools.test.js` — produces TAP output with pass/fail per handler. Zero infrastructure needed.
- **Regression detection:** `node --test tests/*.test.js` — catches regressions across auth, client, and tools layers in one command.
- **Tool count invariant:** `grep -c 'server.tool(' src/tools.js` — must remain 25. Drift means a tool was added/removed without test coverage.
- **Startup health:** `node src/index.js` — prints `🚀 SharePoint Online MCP Server started (stdio)` on success. Any import or wiring error surfaces immediately.
- **Mock call recording:** `MockClient.calls` array records `[methodName, ...args]` for every call, enabling future tests to assert delegation without live services.
- **Error path coverage:** `connect_to_site` and `list_my_sites` handlers return `{ isError: true }` on failure — tests verify this shape explicitly.

## Files Likely Touched

- `tests/tools.test.js`
