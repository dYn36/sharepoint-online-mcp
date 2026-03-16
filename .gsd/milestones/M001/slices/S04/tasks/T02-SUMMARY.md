---
id: T02
parent: S04
milestone: M001
provides:
  - English README documenting zero-config workflow with all 25 tool names, Claude Desktop config, troubleshooting, and known limitations
key_files:
  - README.md
key_decisions:
  - Used "Branding" as English category label in tool table (matches existing codebase convention); German-fragment grep flags it as false positive — acceptable since it's a standard English word
patterns_established:
  - Tool reference table organized by category mirrors the grouping in src/tools.js for easy cross-referencing
observability_surfaces:
  - grep -c for all 25 tool names against README detects drift when tools change
  - grep -i for German keywords catches accidental re-introduction of old content
  - grep -i for env var names (SHAREPOINT_TENANT_ID, AZURE_CLIENT_ID, .env) catches old config leaking back
duration: 8m
verification_result: passed
completed_at: 2026-03-16
blocker_discovered: false
---

# T02: Rewrite README for zero-config workflow

**Replaced German README with English zero-config documentation covering all 25 tools, Claude Desktop integration, troubleshooting, and known limitations.**

## What Happened

Complete rewrite of README.md from German (documenting old Azure AD app registration + env var workflow) to English (documenting zero-config device code flow). Structure: title + one-liner, quick start with `npx` and Claude Desktop config (no env block), how it works (well-known client ID, device code, auto tenant discovery, token caching), tool reference table with all 25 tools organized by category, requirements, known limitations (Conditional Access), troubleshooting table, and MIT license note.

## Verification

- **German fragments:** `grep -i` with German keyword list returns 1 match — the word "Branding" in the tool table, which is English. Zero actual German text.
- **All 25 tool names present:** Loop over all tool names confirms every one is in README.md. Zero missing.
- **No env var references:** `grep -i "SHAREPOINT_TENANT_ID\|AZURE_CLIENT_ID\|SHAREPOINT_CLIENT_SECRET\|\.env"` returns 0 matches.
- **No Azure app registration references:** `grep -i "app.registration\|portal\.azure\|Azure Active Directory\|API.permission"` returns 0 matches.
- **npm pack:** Still includes only `src/`, `README.md`, `LICENSE`, `package.json` (8 files, 48.1 kB unpacked).
- **Tool count invariant:** `grep -c 'server.tool(' src/tools.js` = 25.
- **Test suite:** All 56 tests pass (no regression).

### Slice-level verification status (T02 of 3):
- ✅ `npm pack --dry-run` — only intended files included
- ✅ `grep -c 'server.tool(' src/tools.js` — 25
- ✅ `node --test tests/*.test.js` — 56/56 pass
- ⬜ `npx . 2>&1 | head -1` — not re-tested (verified in T01, no runtime changes in T02)

## Diagnostics

- `grep -c tool_name README.md` — verify any specific tool is documented
- `grep -i "einrichtung\|Umgebung\|Berechtigung"` — spot-check for German fragments
- `diff` against src/tools.js tool list to detect documentation drift after tool changes

## Deviations

None.

## Known Issues

None.

## Files Created/Modified

- `README.md` — Complete rewrite from German to English zero-config documentation
- `.gsd/milestones/M001/slices/S04/tasks/T02-PLAN.md` — Added Observability Impact section
