---
id: S04
milestone: M001
status: ready
---

# S04: npm Package & Polish — Context

<!-- Slice-scoped context. Milestone-only sections (acceptance criteria, completion class,
     milestone sequence) do not belong here — those live in the milestone context. -->

## Goal

`npx sharepoint-online-mcp` installs and runs cleanly, comprehensive English README documents the zero-config workflow with Claude Desktop and Claude CLI setup, error messages are actionable for common failure modes, and the package is publish-ready on npm.

## Why this Slice

S01–S03 deliver a working server with auth, site discovery, and validated tools — but the package isn't distributable yet. package.json has the wrong name, no dependencies listed, no license, and the README is in German documenting the old Azure AD registration flow. This slice makes the package ready for users to install and use with zero friction. It's the terminal slice — nothing depends on it.

## Scope

### In Scope

- **package.json overhaul** — name `sharepoint-online-mcp`, correct bin entry, all dependencies, engines field (Node.js >= 18), keywords, description in English, repository, license field
- **Comprehensive English README** — what it does, install/run (`npx`), how auth works (device code flow), available tool categories, Claude Desktop JSON config example, Claude CLI setup (`claude mcp add`), known limitations (Conditional Access), troubleshooting
- **Claude Desktop config example** — JSON snippet users paste into their MCP config
- **Claude CLI config example** — `claude mcp add` command and equivalent config
- **Mention compatibility** with other MCP clients (Cursor, Windsurf, etc.) without providing specific configs
- **MIT LICENSE file**
- **.npmignore** — exclude .gsd/, tests, dev files from published package
- **Actionable error messages** for top 5–7 failure modes: Conditional Access blocking device code, token expired/revoked, site not found, permission denied, network offline, rate limited, Graph Beta unavailable. Each with one-line explanation + suggested next step.
- **Verify `npx` works locally** — full end-to-end: `npx .` starts server, tools register, auth triggers on first call
- **Prepare only — do not publish to npm.** User publishes manually when ready.

### Out of Scope

- Actually publishing to npm (user does this manually)
- Changelog / CHANGELOG.md
- CI/CD pipeline, GitHub Actions
- Contributing guide
- Config examples for non-Claude MCP clients (beyond a mention of compatibility)
- New features or tool changes — this slice is packaging and polish only
- Version strategy beyond setting initial version in package.json

## Constraints

- Package name is `sharepoint-online-mcp` (D005)
- Plain JavaScript ESM, no build step (D006) — npx must work directly without compilation
- All user-facing strings in English (established in S01/S03 context)
- README must be self-contained — a user should be able to go from zero to working with only the README
- The shebang line (`#!/usr/bin/env node`) must be present in the entry point for bin to work

## Integration Points

### Consumes

- Everything from S01–S03: working auth engine, site discovery tools, validated tool set, English tool descriptions, get_status tool, error handling in client layer
- Existing `src/index.js` — entry point, needs shebang line
- Existing `package.json` — needs complete overhaul

### Produces

- `package.json` — publish-ready with correct name, bin, dependencies, engines, keywords, license, repository, description
- `README.md` — comprehensive English documentation with Claude Desktop + Claude CLI setup
- `LICENSE` — MIT license file
- `.npmignore` — excludes dev/test/gsd files from published package
- Consistent error message layer wrapping common API failure modes with actionable guidance
- Locally verified `npx .` workflow

## Open Questions

- What version number to start at? — Current thinking: `1.0.0` since this is a complete refactor with a new package name. Could also argue `0.1.0` if we want semver room for breaking changes.
- Should `auth-cli.js` be kept as a separate script or removed? — Current thinking: keep it as a debug/test utility, exclude from npm package via .npmignore.
- Does the error message layer belong in a separate file (e.g. `src/errors.js`) or inline in client.js? — Current thinking: separate file for maintainability, decided during planning.
