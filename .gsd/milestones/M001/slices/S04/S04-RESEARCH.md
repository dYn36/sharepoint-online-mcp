# S04: npm Package & Polish — Research

**Date:** 2026-03-16

## Summary

S04 is straightforward polish work. The codebase is functionally complete (S01–S03 delivered zero-config auth, site tools, and all 25 tool validations). What remains is making it npm-publishable and user-friendly.

Three concrete gaps:

1. **package.json is incomplete** — missing `keywords`, `license`, `repository`, `engines`, `files`. Has the correct `name` and `bin` already.
2. **README is entirely in German and documents the old env-var workflow** — needs a full rewrite for zero-config. Steps for Azure app registration, env vars, etc. are now wrong/irrelevant.
3. **Error messages are raw passthroughs** — auth errors (AADSTS codes), Graph 403s, and SP REST failures pass through Microsoft's raw error text with no actionable guidance. No special handling for Conditional Access blocks, consent failures, or scope issues.

A minor fourth item: `auth-cli.js` needs a shebang for standalone execution and should be in `bin` or documented.

## Recommendation

Three independent tasks, parallelizable:

1. **package.json polish** — add missing npm metadata fields, add `files` whitelist, add `.npmignore`, verify `npx sharepoint-online-mcp` works locally via `npm link` or `npx .`
2. **README rewrite** — English, zero-config workflow, Claude Desktop config (no env vars), tool reference table, known limitations (Conditional Access), troubleshooting
3. **Error message enhancement** — wrap auth and client error paths to detect common AADSTS codes and HTTP status codes, provide actionable guidance

Tasks 1 and 2 have no code overlap. Task 3 touches `src/auth.js` and `src/client.js` which are stable (no changes since S01). All three can be planned in parallel.

## Implementation Landscape

### Key Files

- `package.json` — needs `keywords`, `license`, `repository`, `engines`, `files` fields. `name`, `version`, `description`, `bin`, `main`, `dependencies` are already correct.
- `README.md` — full rewrite. Currently German + old Azure app registration flow. Replace with English zero-config docs.
- `src/auth.js` — error paths at lines 44 (tenant discovery network error), 141 (silent acquisition failure), and the device code flow call need actionable error wrapping. Currently passes raw MSAL errors.
- `src/client.js` — error paths at lines 29 (Graph `API {status}: {body}`) and 63 (SP REST `SP REST {status}: {body}`) need to detect 401/403/404 and add user-facing guidance.
- `src/index.js` — startup error handler at line 37 is fine but could add a hint about Node version requirements.
- `.gitignore` — already exists and is adequate. No `.npmignore` exists — need one to exclude tests, .gsd, etc. from the published package.

### Build Order

All three tasks are independent:
1. **Package metadata + npx verification** — fastest to verify (npm pack, check contents, npm link test)
2. **README** — pure content, no code changes
3. **Error messages** — touches source, needs test updates

Verification of the npx flow is the slice's key proof — do that in task 1 and re-verify after task 3 changes.

### Verification Approach

- `npm pack --dry-run` — verify only intended files are included (src/, package.json, README.md, LICENSE)
- `npx .` from project root — verify bin entry starts the server
- `node --test tests/*.test.js` — full regression (56 tests) after error message changes
- `grep -c 'server.tool(' src/tools.js` = 25 — tool count invariant
- Manual review of README content for accuracy against actual behavior

## Constraints

- No build step — plain ESM JavaScript. Package must work directly via `npx` without compilation.
- `src/index.js` already has `#!/usr/bin/env node` shebang — confirmed correct.
- `files` field in package.json must whitelist only what npm needs: `src/`, `README.md`, `LICENSE`. Tests and .gsd must be excluded.
- Node.js >= 18 (global fetch used in auth.js tenant discovery, MSAL Node compatibility).

## Common Pitfalls

- **`files` vs `.npmignore` precedence** — `files` is a whitelist and takes precedence over `.npmignore`. Use `files` as the primary mechanism; `.npmignore` is belt-and-suspenders for anything `files` might miss.
- **bin entry needs executable permission** — `src/index.js` must be `chmod +x` for `npx` to work on unix. Check current permissions.
- **README code examples must match actual tool names** — the tool names were translated to English in S03. README examples must use the actual registered tool names from `src/tools.js`.

## Open Risks

- `npm link` / `npx .` testing on the worktree may not perfectly simulate a real npm install (symlinks vs actual install). `npm pack` + `npm install <tarball>` is the authoritative test but heavier.
- License choice needs user input if not already decided — MIT is standard for this type of project but wasn't specified.
