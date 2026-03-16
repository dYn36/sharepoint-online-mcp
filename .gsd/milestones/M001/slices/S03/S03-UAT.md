# S03: Tool Validation & Dual-Audience Tokens — UAT

**Milestone:** M001
**Written:** 2026-03-16

## UAT Type

- UAT mode: artifact-driven
- Why this mode is sufficient: S03 is a test-only slice — all deliverables are mock-based unit tests proving contract compliance. No new runtime behavior to verify interactively.

## Preconditions

- Node.js ≥ 18 installed
- Working directory is the project root (or M001 worktree)
- `npm install` has been run (dependencies available in `node_modules/`)

## Smoke Test

Run `node --test tests/tools.test.js` — should output 38 passing tests with 0 failures in under 1 second.

## Test Cases

### 1. Full test suite passes with zero failures

1. Run `node --test tests/*.test.js`
2. **Expected:** 56 tests pass (38 tools + 12 auth + 6 client), 0 failures, 0 skipped

### 2. Tool count invariant is maintained

1. Run `grep -c 'server.tool(' src/tools.js`
2. **Expected:** Output is exactly `25`

### 3. Server starts without import breakage

1. Run `node src/index.js` (kill after startup message appears on stderr)
2. **Expected:** Outputs "🚀 SharePoint Online MCP Server started (stdio)" with no errors

### 4. Every tool category has test coverage

1. Run `node --test tests/tools.test.js 2>&1 | grep '✔'`
2. **Expected:** All these suite names appear as passing:
   - Sites (5 tests)
   - Pages (5+ tests)
   - Layout (3 tests)
   - Web Parts (6 tests)
   - Navigation (4 tests)
   - Branding (2 tests)
   - Auth & Utility (4 tests)
   - Additional page tools (2 tests)

### 5. Error paths return isError shape

1. Run `node --test tests/tools.test.js --test-name-pattern="error|isError"` 
2. **Expected:** At least 2 tests pass — connect_to_site error and list_my_sites error, both verifying `isError: true` in response

### 6. MCP content shape is validated across all handlers

1. In `tests/tools.test.js`, search for `assertMcpContent` usage
2. **Expected:** Every handler test calls `assertMcpContent(result)` which asserts `result.content` is an array with at least one `{ type: "text", text: <string> }` entry

### 7. Mock infrastructure is reusable

1. Review MockServer, MockClient, MockAuth classes in `tests/tools.test.js`
2. **Expected:**
   - MockServer exposes `call(name, args)` to invoke any registered handler
   - MockClient has stub methods for all 20 client methods, each recording calls in `this.calls`
   - MockAuth tracks `logout()` via boolean flag
   - All three are instantiated fresh per test via `beforeEach`

## Edge Cases

### Graph vs SP REST delegation is distinguishable

1. Check navigation handler tests (get_navigation with 'quick', get_navigation with 'top', add_navigation_link, remove_navigation_link)
2. **Expected:** These call `getNavigation`, `getTopNavigation`, `addNavigationNode`, `deleteNavigationNode` on the client — these are SP REST methods, not Graph methods. The mock records confirm the right method family is called.

### Base64 decode in upload_asset

1. Check the upload_asset test in the Branding suite
2. **Expected:** Test passes a base64-encoded string as `fileContent`, and the handler decodes it to a Buffer before calling `uploadFile`. The test asserts the third argument is a Buffer instance.

### autoPublish flag on create_page

1. Check the create_page+autoPublish test in the Pages suite
2. **Expected:** When `autoPublish: true` is passed, the handler calls both `createPage` and then `publishPage`. MockClient.calls records both calls in order.

## Failure Signals

- Any test failure in `node --test tests/*.test.js` — indicates a regression or broken contract
- `grep -c 'server.tool(' src/tools.js` returning anything other than 25 — tool was added/removed without test update
- `node src/index.js` failing to start — import or wiring error introduced

## Requirements Proved By This UAT

- R008 (Page CRUD) — contract-level proof via mock delegation tests for all page tools
- R009 (Canvas Layout) — contract-level proof via mock delegation tests for section/layout tools
- R010 (Web Parts) — contract-level proof for all 6 web part types
- R011 (Navigation) — contract-level proof that navigation tools call SP REST client methods
- R012 (Branding) — contract-level proof for logo and asset upload delegation

## Not Proven By This UAT

- Live dual-audience token acquisition (Graph vs SP REST tokens against real Azure AD)
- Actual SharePoint API responses matching mock return shapes
- End-to-end: start server → authenticate → call tool → see SharePoint change
- These are deferred to milestone-level UAT after S04

## Notes for Tester

- All tests run locally with zero network calls — safe to run offline
- Tests complete in under 1 second total — no timeouts expected
- The `overrides` injection in `registerTools()` is only used by tests — production code path passes no overrides and uses the real `discoverTenantId` import
