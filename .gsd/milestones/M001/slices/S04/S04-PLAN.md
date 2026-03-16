# S04: npm Package & Polish

**Goal:** `npx sharepoint-online-mcp` installs and runs cleanly, README documents zero-config workflow, error messages are actionable, package is publish-ready.
**Demo:** Run `npm pack`, verify only intended files are included. Run `npx .` and confirm the server starts. README is in English and documents the zero-config flow. Common auth/API errors produce user-facing guidance.

## Must-Haves

- `package.json` has `keywords`, `license`, `repository`, `engines`, `files` fields
- `files` whitelist includes only `src/`, `README.md`, `LICENSE` — no tests, no `.gsd/`
- `.npmignore` as belt-and-suspenders exclusion
- `src/index.js` has executable permission (`chmod +x`)
- `LICENSE` file exists (MIT)
- `npx .` starts the server successfully
- README is in English, documents zero-config workflow with Claude Desktop config
- README includes tool reference table with all 25 actual tool names
- README documents known limitations (Conditional Access) and troubleshooting
- Auth errors (AADSTS codes, network failures) produce actionable guidance
- Client errors (401, 403, 404) produce user-facing messages explaining what to do
- All 56 existing tests pass after changes

## Proof Level

- This slice proves: final-assembly
- Real runtime required: yes (npx startup, npm pack)
- Human/UAT required: no (live SharePoint testing is milestone-level UAT)

## Verification

- `npm pack --dry-run 2>&1` — verify only `src/`, `README.md`, `LICENSE`, `package.json` are included. No `tests/`, `.gsd/`, `.git/`.
- `npx . 2>&1 | head -1` — outputs "🚀 SharePoint Online MCP Server started (stdio)" (tests bin entry + shebang + executable permission)
- `node --test tests/*.test.js` — all 56 tests pass (regression after error message changes)
- `grep -c 'server.tool(' src/tools.js` — still 25

## Integration Closure

- Upstream surfaces consumed: `src/auth.js` (error paths), `src/client.js` (error paths), `src/tools.js` (tool names for README), `src/index.js` (bin entry)
- New wiring introduced in this slice: none (polish only)
- What remains before the milestone is truly usable end-to-end: live UAT against real SharePoint tenant (milestone-level, not slice-level)

## Tasks

- [ ] **T01: Make package npm-publish-ready** `est:25m`
  - Why: R013 — `npx sharepoint-online-mcp` must install and run cleanly. Currently missing `files`, `engines`, `keywords`, `license`, `repository` in package.json. No LICENSE file. No `.npmignore`. `src/index.js` lacks executable permission.
  - Files: `package.json`, `LICENSE`, `.npmignore`, `src/index.js`
  - Do: Add `files: ["src/", "README.md", "LICENSE"]`, `engines: { node: ">=18" }`, `keywords`, `license: "MIT"`, `repository` to package.json. Create MIT LICENSE file (year 2026). Create `.npmignore` excluding tests/, .gsd/, .git/, node_modules/, .env. Run `chmod +x src/index.js`. Verify with `npm pack --dry-run` and `npx .`.
  - Verify: `npm pack --dry-run` shows only intended files. `npx . 2>&1 | head -1` shows startup message. `node --test tests/*.test.js` still passes all 56 tests.
  - Done when: `npm pack --dry-run` includes only `src/`, `README.md`, `LICENSE`, `package.json` and `npx .` starts the server.

- [ ] **T02: Rewrite README for zero-config workflow** `est:30m`
  - Why: R013 (documentation for npm package), partially R015 (troubleshooting). README is entirely in German and documents the old Azure app registration + env var workflow, which is now wrong. Users need English docs showing the zero-config flow.
  - Files: `README.md`
  - Do: Full rewrite in English. Structure: title + one-liner, quick start (`npx sharepoint-online-mcp` + Claude Desktop JSON config with no env vars), how it works (device code flow, auto tenant discovery, token caching), tool reference table (all 25 tools organized by category with actual names from `src/tools.js`: `search_sites`, `get_site_details`, `get_site_by_url`, `connect_to_site`, `list_my_sites`, `list_pages`, `get_page`, `create_page`, `update_page`, `publish_page`, `delete_page`, `add_section`, `get_page_layout`, `add_text_webpart`, `add_image_webpart`, `add_spacer`, `add_divider`, `add_custom_webpart`, `get_navigation`, `add_navigation_link`, `remove_navigation_link`, `set_site_logo`, `upload_asset`, `get_design_templates`, `disconnect`), known limitations (Conditional Access blocking, requires Node >= 18), troubleshooting section (common errors + what to do). No Azure app registration steps — the whole point is zero-config.
  - Verify: Manual review — README is English, no German fragments, no references to env vars or Azure app registration, all 25 tool names match `src/tools.js`, Claude Desktop config has no env vars.
  - Done when: README accurately documents the zero-config workflow with correct tool names and Claude Desktop configuration.

- [ ] **T03: Wrap auth and client errors with actionable guidance** `est:30m`
  - Why: R015 — error messages currently pass through raw Microsoft error text (AADSTS codes, HTTP status codes with JSON bodies). Users in a zero-config setup have no config knobs to debug with, so error messages must explain what went wrong and what to try.
  - Files: `src/auth.js`, `src/client.js`, `tests/auth.test.js`, `tests/client.test.js`
  - Do: In `src/auth.js` — wrap the device code flow error path to detect common AADSTS codes: `AADSTS50076`/`AADSTS53003` (Conditional Access — "Your organization blocks device code flow. Contact your IT admin or try from a different network."), `AADSTS700016` (app not found — should not happen with well-known client ID but defensive), `AADSTS50059` (no tenant found — "Could not find tenant. Check that the SharePoint URL is correct."), generic auth failure ("Authentication failed. Try running `disconnect` and authenticating again."). In `src/client.js` — enhance the two `throw new Error` lines: 401 → "Authentication expired or revoked. Use the `disconnect` tool and reconnect.", 403 → "Access denied. Your account may not have permission for this operation. Check site permissions.", 404 → "Resource not found. Verify the site URL and resource ID are correct.", keep existing format for other status codes but prefix with "SharePoint API error: ". Add tests for the new error message patterns — test that specific status codes produce the expected guidance strings.
  - Verify: `node --test tests/*.test.js` — all tests pass including new error message tests. `grep -c 'server.tool(' src/tools.js` still equals 25 (no tool changes).
  - Done when: Auth errors with known AADSTS codes and client errors with 401/403/404 produce actionable guidance messages. All tests pass.

## Files Likely Touched

- `package.json`
- `LICENSE`
- `.npmignore`
- `src/index.js` (permissions only)
- `README.md`
- `src/auth.js`
- `src/client.js`
- `tests/auth.test.js`
- `tests/client.test.js`
