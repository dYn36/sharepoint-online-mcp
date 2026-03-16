---
id: S03
parent: M001
milestone: M001
provides:
  - 31 new tool handler tests covering all 25 MCP tools across every category (sites, pages, layout, web parts, navigation, branding, auth/utility)
  - Mock infrastructure (MockServer, MockClient, MockAuth) reusable for future tool testing
  - Contract proof that Graph-based tools and SP REST tools delegate correctly through the client layer
  - Injection seam via overrides parameter on registerTools() for dependency substitution in tests
requires:
  - slice: S01
    provides: SharePointAuth with getAccessToken(resource), discoverTenantId
  - slice: S02
    provides: connect_to_site, list_my_sites, search_sites tools registered in registerTools()
affects:
  - S04
key_files:
  - tests/tools.test.js
  - src/tools.js
key_decisions:
  - Added overrides parameter to registerTools instead of module-level mocking — minimal source change, avoids experimental Node flags
patterns_established:
  - MockServer.call(name, args) pattern for invoking registered tool handlers in tests
  - MockClient._record() call-tracking pattern for asserting delegation arguments
  - assertMcpContent() helper for validating MCP response shape across all tests
observability_surfaces:
  - node --test tests/tools.test.js — TAP output with per-handler pass/fail
  - node --test tests/*.test.js — full regression across auth, client, and tools (56 tests)
  - MockClient.calls array records [methodName, ...args] for delegation assertion
drill_down_paths:
  - .gsd/milestones/M001/slices/S03/tasks/T01-SUMMARY.md
duration: 15m
verification_result: passed
completed_at: 2026-03-16
---

# S03: Tool Validation & Dual-Audience Tokens

**31 handler tests prove all 25 MCP tools delegate correctly through the client layer — Graph API tools route to Graph audience, SP REST tools route to SharePoint audience, error paths return isError:true with actionable messages.**

## What Happened

Built mock infrastructure and a comprehensive test suite in a single task. The approach: create MockServer (captures `tool()` registrations, exposes `call(name, args)` to invoke handlers directly), MockClient (20 stub methods matching every client method, each recording `[method, ...args]` and returning canned data), and MockAuth (tracks `logout()` calls). A 2-line source change to `registerTools()` added an `overrides` parameter for injecting `discoverTenantId` in tests without experimental module mocking.

31 new handler tests cover every tool category:
- **Sites (5):** search_sites, get_site_details, get_site_by_url, connect_to_site, list_my_sites
- **Pages (7):** list_pages, get_page, create_page, create_page+autoPublish, delete_page, update_page, publish_page
- **Layout (3):** add_section, add_section+emphasis, get_page_layout
- **Web Parts (6):** text, image, spacer, divider, custom webpart, position shape
- **Navigation (4):** get quick, get top, add link, remove link
- **Branding (2):** set_site_logo, upload_asset (with base64 decode verification)
- **Auth/Utility (4):** disconnect, get_design_templates, connect_to_site error, list_my_sites error

Every test verifies three things: (1) correct client method called with expected args, (2) MCP content shape `{ content: [{ type: "text", text }] }`, (3) error paths return `isError: true`.

## Verification

| Check | Result |
|-------|--------|
| `node --test tests/tools.test.js` | ✅ 38 pass, 0 fail (7 URL parsing + 31 handler) |
| `node --test tests/auth.test.js tests/client.test.js` | ✅ 18 pass, 0 fail |
| `grep -c 'server.tool(' src/tools.js` | ✅ 25 (tool count invariant held) |
| `node src/index.js` startup | ✅ "🚀 SharePoint Online MCP Server started (stdio)" |
| Total tests | 56 pass, 0 fail |

## Requirements Advanced

- R008 (Page CRUD) — All page handlers (list, get, create, update, delete, publish) tested for correct delegation and MCP response shape
- R009 (Canvas Layout Editing) — add_section and get_page_layout handlers tested with template and emphasis parameters
- R010 (Web Part Management) — All 6 web part types tested: text, image, spacer, divider, custom, plus position shape validation
- R011 (Navigation Management) — Quick Launch and Top Navigation get/add/remove handlers tested with correct SP REST delegation
- R012 (Branding) — set_site_logo and upload_asset handlers tested, including base64 buffer decode verification
- R015 (Clear Error Messages) — Error paths on connect_to_site and list_my_sites return `isError: true` with messages

## Requirements Validated

- R008 — Contract-level: mock tests prove every page tool calls the correct client method with correct args and returns valid MCP content
- R009 — Contract-level: layout tools delegate correctly with template/emphasis parameters
- R010 — Contract-level: all web part types delegate with correct data shapes
- R011 — Contract-level: navigation tools call SP REST client methods (not Graph), proving audience routing
- R012 — Contract-level: branding tools delegate correctly, upload_asset decodes base64 to Buffer

## New Requirements Surfaced

- None

## Requirements Invalidated or Re-scoped

- None

## Deviations

- Delivered 31 new tests vs plan's ~30 target — added update_page and publish_page tests as trivial handlers that deserved coverage.

## Known Limitations

- All validation is contract-level (mock-based). Live dual-audience token behavior against real SharePoint is not proven in this slice — deferred to milestone-level UAT.
- The `overrides` parameter on `registerTools()` is a test-only seam. It's not validated for production use patterns.

## Follow-ups

- S04 should include a live smoke test as part of milestone UAT — real device code auth, real Graph + SP REST calls.
- The MockClient infrastructure can be extended when new tools are added post-M001.

## Files Created/Modified

- `src/tools.js` — Added `overrides = {}` parameter to `registerTools` and `_discoverTenantId` injection (2 lines changed)
- `tests/tools.test.js` — Expanded from 39 lines to ~340 lines with MockServer/MockClient/MockAuth infrastructure and 31 new handler tests

## Forward Intelligence

### What the next slice should know
- All 25 tools are tested at contract level. The test suite is the fastest way to check for regressions after any source change: `node --test tests/*.test.js` runs all 56 tests in under 1 second.
- The `overrides` parameter on `registerTools()` is the injection seam for testing — it currently only substitutes `discoverTenantId`, but the pattern extends to any module-level import.

### What's fragile
- `grep -c 'server.tool(' src/tools.js` = 25 is used as a tool count invariant. If S04 adds or removes tools, this check and test coverage must be updated in lockstep.
- MockClient stubs return minimal canned data. If tool handlers start inspecting response structure more deeply, stubs may need richer return values.

### Authoritative diagnostics
- `node --test tests/*.test.js` — single command, full regression, TAP output. Trust this over any partial check.
- `MockClient.calls` array in tests — shows exactly which client method was called with which args. First place to look when debugging delegation.

### What assumptions changed
- Plan estimated ~30 tests needed — actual was 31, with 2 extra page tool tests. No significant assumption changes.
- Plan said "no source changes" — reality required a minimal 2-line change to `src/tools.js` for the overrides injection seam. Correct tradeoff.
