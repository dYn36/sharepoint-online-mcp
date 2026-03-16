# S04: npm Package & Polish — UAT

**Milestone:** M001
**Written:** 2026-03-16

## UAT Type

- UAT mode: artifact-driven
- Why this mode is sufficient: S04 is packaging and polish — npm pack output, README content, and error message strings are all verifiable from artifacts and test output without a live SharePoint tenant.

## Preconditions

- Node.js >= 18 installed
- Working directory is the project root (where `package.json` lives)
- `npm install` has been run (dependencies present in `node_modules/`)

## Smoke Test

Run `node src/index.js` — should emit `🚀 SharePoint Online MCP Server started (stdio)` on stderr within 1 second. If it does, bin entry, shebang, module resolution, and executable permission all work.

## Test Cases

### 1. npm pack produces clean artifact

1. Run `npm pack --dry-run 2>&1`
2. Count the listed files
3. **Expected:** Exactly 8 files: `LICENSE`, `README.md`, `package.json`, `src/auth-cli.js`, `src/auth.js`, `src/client.js`, `src/index.js`, `src/tools.js`
4. Verify no files from `tests/`, `.gsd/`, `.git/`, `.env`, or `node_modules/` appear in the list
5. **Expected:** Zero excluded files present

### 2. Package metadata is complete

1. Run `node -e "const p=JSON.parse(require('fs').readFileSync('package.json'));console.log(JSON.stringify({license:p.license,engines:p.engines,files:p.files,bin:p.bin,keywords:p.keywords,repository:p.repository},null,2))"`
2. **Expected:**
   - `license` is `"MIT"`
   - `engines.node` is `">=18"`
   - `files` includes `"src/"`, `"README.md"`, `"LICENSE"`
   - `bin` maps `"sharepoint-online-mcp"` to `"src/index.js"`
   - `keywords` array is non-empty
   - `repository` field is present

### 3. Server starts without env vars

1. Unset all SharePoint/Azure env vars: `unset SHAREPOINT_TENANT_ID AZURE_CLIENT_ID SHAREPOINT_CLIENT_SECRET`
2. Run `node src/index.js &` and capture first stderr line
3. **Expected:** `🚀 SharePoint Online MCP Server started (stdio)`
4. Kill the background process

### 4. README is English with zero-config workflow

1. Open `README.md`
2. Verify the quick start section shows `npx sharepoint-online-mcp` (no env vars)
3. Verify Claude Desktop config JSON has no `env` block
4. Run `grep -i "SHAREPOINT_TENANT_ID\|AZURE_CLIENT_ID\|SHAREPOINT_CLIENT_SECRET\|\.env" README.md`
5. **Expected:** Zero matches
6. Run `grep -i "app.registration\|portal\.azure\|Azure Active Directory\|API.permission" README.md`
7. **Expected:** Zero matches
8. Scan for German fragments: `grep -iE "einrichtung|Umgebung|Berechtigung|Konfiguration|Voraussetzung" README.md`
9. **Expected:** Zero matches

### 5. All 25 tool names documented in README

1. For each tool name (`search_sites`, `get_site_details`, `get_site_by_url`, `connect_to_site`, `list_my_sites`, `list_pages`, `get_page`, `create_page`, `update_page`, `publish_page`, `delete_page`, `add_section`, `get_page_layout`, `add_text_webpart`, `add_image_webpart`, `add_spacer`, `add_divider`, `add_custom_webpart`, `get_navigation`, `add_navigation_link`, `remove_navigation_link`, `set_site_logo`, `upload_asset`, `get_design_templates`, `disconnect`), run `grep -c '<tool_name>' README.md`
2. **Expected:** Every tool name appears at least once

### 6. Auth error messages are actionable

1. Run `node --test tests/auth.test.js 2>&1`
2. **Expected:** All 8 `wrapAuthError` tests pass, covering:
   - AADSTS50076 → mentions "Conditional Access" and "IT admin"
   - AADSTS53003 → mentions "Conditional Access" and "IT admin"
   - AADSTS700016 → mentions "application"
   - AADSTS50059 → mentions "tenant" and "SharePoint URL"
   - Unknown AADSTS code → preserves code in message
   - Non-AADSTS error → generic auth failure with disconnect hint
   - Error in message property (not errorCode) → still detected
   - Non-Error object → handled gracefully

### 7. Client error messages are actionable

1. Run `node --test tests/client.test.js 2>&1`
2. **Expected:** All 8 HTTP error message tests pass, covering:
   - Graph 401 → mentions "Authentication expired" and "disconnect"
   - Graph 403 → mentions "Access denied" and "permission"
   - Graph 404 → mentions "not found" and "verify"
   - Graph other status → includes status code and "SharePoint API error"
   - SP REST 401 → mentions "Authentication expired" and "disconnect"
   - SP REST 403 → mentions "Access denied" and "permission"
   - SP REST 404 → mentions "not found" and "verify"
   - SP REST other status → includes status code and "SharePoint API error"

### 8. Full test suite passes

1. Run `node --test tests/*.test.js`
2. **Expected:** 72 tests pass, 0 fail, 0 cancelled
3. Run `grep -c 'server.tool(' src/tools.js`
4. **Expected:** 25

## Edge Cases

### MIT LICENSE file present and well-formed

1. Run `cat LICENSE | head -3`
2. **Expected:** Contains "MIT License" and copyright year 2026

### .npmignore excludes sensitive paths

1. Run `cat .npmignore`
2. **Expected:** Contains entries for `tests/`, `.gsd/`, `.git/`, `.env`, `node_modules/`

### src/index.js is executable

1. Run `ls -la src/index.js`
2. **Expected:** File has executable permission (`x` bit set)
3. Run `head -1 src/index.js`
4. **Expected:** Shebang line `#!/usr/bin/env node`

### Conditional Access documented as limitation

1. Run `grep -i "conditional access" README.md`
2. **Expected:** At least one match in a "Known Limitations" or similar section explaining that some tenants block device code flow

## Failure Signals

- `npm pack --dry-run` lists files outside `src/`, `README.md`, `LICENSE`, `package.json` — file leak
- `node src/index.js` fails to start or produces no output — broken bin entry or module resolution
- Any of the 72 tests fail — regression in auth, client, or tools
- `grep -c 'server.tool(' src/tools.js` ≠ 25 — tool count drift
- German text in README — incomplete rewrite
- Env var names in README — old workflow leaking through
- Error messages in tests don't contain actionable keywords — error wrapping broken

## Requirements Proved By This UAT

- R013 — Package installs and runs from `npx sharepoint-online-mcp`. Verified via npm pack, metadata check, and server startup.
- R015 — Error messages are actionable. Verified via 16 unit tests covering AADSTS codes and HTTP status codes.
- R014 — No secrets in published artifact. Verified via npm pack dry-run file list.

## Not Proven By This UAT

- R013 actual npm registry publish — package is publish-ready but not published
- Live SharePoint tenant authentication and tool execution — deferred to milestone-level UAT
- Performance under concurrent tool calls — not in scope for this milestone

## Notes for Tester

- The `npx .` test (test case 3) starts a stdio-based MCP server that reads stdin — it will hang waiting for MCP protocol input. Verify the startup message and then kill the process.
- "Branding" appears in the README tool table as a category name — this is English, not a German fragment.
- The 72-test count (up from 56) is expected — T03 added 16 new tests for error message coverage.
