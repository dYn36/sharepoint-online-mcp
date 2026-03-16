---
id: T01
parent: S03
milestone: M001
provides:
  - Mock infrastructure (MockServer, MockClient, MockAuth) for testing all 25 MCP tool handlers
  - Injection seam via overrides parameter on registerTools for dependency substitution
  - Full test coverage across all tool categories: sites, pages, layout, web parts, navigation, branding, auth
key_files:
  - tests/tools.test.js
  - src/tools.js
key_decisions:
  - Added overrides parameter to registerTools instead of module-level mocking — keeps source change minimal and explicit
patterns_established:
  - MockServer.call(name, args) pattern for invoking registered tool handlers in tests
  - MockClient._record() call-tracking pattern for asserting delegation arguments
  - assertMcpContent() helper for validating MCP response shape across all tests
observability_surfaces:
  - node --test tests/tools.test.js — TAP output with per-handler pass/fail
  - MockClient.calls array records [methodName, ...args] for delegation assertion
duration: 15m
verification_result: passed
completed_at: 2026-03-16
blocker_discovered: false
---

# T01: Tool handler test suite with mock server/client/auth

**Built mock infrastructure and 31 new tests covering all 25 MCP tool handlers across every category — sites, pages, layout, web parts, navigation, branding, auth/utility — with error path coverage.**

## What Happened

Added an `overrides = {}` 4th parameter to `registerTools()` in `src/tools.js` (2 lines changed) to allow injecting `discoverTenantId` for the `connect_to_site` handler without module-level mocking.

Built three mock classes in `tests/tools.test.js`:
- **MockServer** — captures `tool()` registrations, exposes `call(name, args)` to invoke handlers directly
- **MockClient** — 20 stub methods matching every client method used by tools, each recording `[method, ...args]` in `this.calls` array and returning minimal canned data
- **MockAuth** — tracks `logout()` calls via boolean flag

Wrote 31 new handler tests (38 total with 7 existing URL parsing tests):
- Sites (5): search_sites, get_site_details, get_site_by_url, connect_to_site, list_my_sites
- Pages (5): list_pages, get_page, create_page, create_page+autoPublish, delete_page
- Layout (3): add_section, add_section+emphasis, get_page_layout
- Web Parts (6): text, image, spacer, divider, custom, position shape
- Navigation (4): get quick, get top, add link, remove link
- Branding (2): set_site_logo, upload_asset (with base64 decode verification)
- Auth/Utility (4): disconnect, get_design_templates, connect_to_site error, list_my_sites error
- Additional pages (2): update_page, publish_page

Every test verifies: (1) correct client method called with expected args, (2) MCP content shape `{ content: [{ type: "text", text }] }`, (3) error paths return `isError: true`.

## Verification

All four slice-level checks pass:

| Check | Result |
|-------|--------|
| `node --test tests/tools.test.js` | ✅ 38 pass, 0 fail (7 URL + 31 handler) |
| `node --test tests/auth.test.js tests/client.test.js` | ✅ 18 pass, 0 fail |
| `grep -c 'server.tool(' src/tools.js` | ✅ 25 |
| `node src/index.js` | ✅ Starts with "🚀 SharePoint Online MCP Server started (stdio)" |

Total across all test files: 56 tests pass (38 + 18), exceeding the 50+ slice target.

## Diagnostics

- Run `node --test tests/tools.test.js` for per-handler pass/fail TAP output
- MockClient.calls array pattern available for future tests to inspect delegation args
- Error shape `{ isError: true, content: [...] }` is now a tested contract

## Deviations

- Added 2 extra tests (update_page, publish_page) beyond the plan's ~29 minimum — these were trivial handlers that deserved coverage
- Plan said "~30 new tests" — delivered 31 new (38 total in file)

## Known Issues

None.

## Files Created/Modified

- `src/tools.js` — Added `overrides = {}` parameter to `registerTools` and `_discoverTenantId` injection (2 lines)
- `tests/tools.test.js` — Expanded from 39 lines to ~340 lines with MockServer/MockClient/MockAuth infrastructure and 31 new handler tests
