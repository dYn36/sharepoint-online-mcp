# Project

## What This Is

An MCP (Model Context Protocol) server that lets AI assistants design and edit SharePoint Online sites — pages, layouts, web parts, navigation, branding — using the authenticated user's own permissions. Zero-config: no Azure Portal, no app registration, no env vars.

## Core Value

A user runs `npx sharepoint-online-mcp` with zero configuration, authenticates via device code in their browser, and Claude can immediately create/edit/publish SharePoint pages on sites they have access to.

## Current State

**M001 (Zero-Config SharePoint MCP) — complete.** All four slices delivered. The server starts with zero env vars, authenticates via device code using Microsoft Office's well-known client ID, auto-discovers tenant from SharePoint URLs, and routes dual-audience tokens to 25 MCP tools (pages, layout, web parts, navigation, branding, site discovery). Package is npm-publish-ready: `npm pack` produces a clean 8-file, 12.1kB artifact. 72 unit tests passing across 3 test files. All 15 in-scope requirements validated at contract level. Human UAT against a live SharePoint tenant is the remaining step before `npm publish`.

## Architecture / Key Patterns

- **Runtime:** Node.js ESM, plain JavaScript (no TypeScript, no build step)
- **Protocol:** MCP over stdio (`@modelcontextprotocol/sdk`)
- **Auth:** MSAL Node (`@azure/msal-node` ^5.1.0) with Device Code Flow, well-known client ID `d3590ed6-52b3-4102-aeff-aad2292ab01c`
- **API:** Microsoft Graph v1.0 + Beta for pages/layout, SharePoint REST API for navigation/branding
- **Structure:**
  - `src/auth.js` — MSAL wrapper, token cache, device code flow, tenant discovery, wrapAuthError
  - `src/client.js` — Graph API + SP REST API client with per-resource token routing, injectable fetchFn
  - `src/tools.js` — 25 MCP tool definitions (zod schemas + handlers), parseSharePointUrl export
  - `src/index.js` — Server entrypoint, stdio transport
  - `src/auth-cli.js` — Standalone auth test CLI

## Capability Contract

See `.gsd/REQUIREMENTS.md` for the explicit capability contract. All 15 in-scope requirements are validated. R016 (Conditional Access fallback) is deferred. R017-R018 are out of scope.

## Milestone Sequence

- [x] M001: Zero-Config SharePoint MCP — Complete. All slices delivered (S01–S04). 72 tests passing. Package publish-ready. Human UAT pending.
