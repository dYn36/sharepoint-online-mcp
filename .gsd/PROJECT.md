# Project

## What This Is

An MCP (Model Context Protocol) server that lets AI assistants design and edit SharePoint Online sites — pages, layouts, web parts, navigation, branding — using the authenticated user's own permissions. Zero-config: no Azure Portal, no app registration, no env vars.

## Core Value

A user runs `npx sharepoint-online-mcp` with zero configuration, authenticates via device code in their browser, and Claude can immediately create/edit/publish SharePoint pages on sites they have access to.

## Current State

All four slices of M001 complete. Package is publish-ready: `npm pack` produces a clean 8-file, 12.1kB artifact. Server starts with `node src/index.js` — no env vars, no config files. Uses Microsoft Office well-known client ID for device code flow authentication. Token caching to `~/.sharepoint-mcp-cache.json` for cross-restart persistence. Dual-audience token routing: Graph API and SP REST API get per-resource tokens. 25 MCP tools registered, all with English descriptions. English README documents zero-config workflow with Claude Desktop config, all 25 tools, troubleshooting, and known limitations. Auth and API errors produce actionable guidance messages. 72 unit tests passing across 3 test files. Milestone-level live UAT against a real SharePoint tenant is the remaining validation step before npm publish.

## Architecture / Key Patterns

- **Runtime:** Node.js ESM, plain JavaScript (no TypeScript, no build step)
- **Protocol:** MCP over stdio (`@modelcontextprotocol/sdk`)
- **Auth:** MSAL Node (`@azure/msal-node`) with Device Code Flow
- **API:** Microsoft Graph v1.0 + Beta for pages/layout, SharePoint REST API for navigation/branding
- **Structure:**
  - `src/auth.js` — MSAL wrapper, token cache, device code flow, tenant discovery, error wrapping (wrapAuthError)
  - `src/client.js` — Graph API + SP REST API client with per-resource token routing, injectable fetchFn
  - `src/tools.js` — 25 MCP tool definitions (zod schemas + handlers), `parseSharePointUrl` export
  - `src/index.js` — Server entrypoint, stdio transport
  - `src/auth-cli.js` — Standalone auth test CLI

## Capability Contract

See `.gsd/REQUIREMENTS.md` for the explicit capability contract, requirement status, and coverage mapping.

## Milestone Sequence

- [ ] M001: Zero-Config SharePoint MCP — All slices complete (S01–S04). Pending: milestone-level live UAT, npm publish.
