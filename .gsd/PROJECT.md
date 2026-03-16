# Project

## What This Is

An MCP (Model Context Protocol) server that lets AI assistants design and edit SharePoint Online sites — pages, layouts, web parts, navigation, branding — using the authenticated user's own permissions. Zero-config: no Azure Portal, no app registration, no env vars.

## Core Value

A user runs `npx sharepoint-online-mcp` with zero configuration, authenticates via device code in their browser, and Claude can immediately create/edit/publish SharePoint pages on sites they have access to.

## Current State

Auth layer refactored to zero-config. Server starts with `node src/index.js` — no env vars, no config files. Uses Microsoft Office well-known client ID for device code flow authentication. Token caching to `~/.sharepoint-mcp-cache.json` for cross-restart persistence. Dual-audience token routing: Graph API and SP REST API get per-resource tokens. 25 MCP tools registered, all with English descriptions. Site discovery tools added: `connect_to_site` (URL → tenant discovery → site resolution), `list_my_sites` (followed sites). 25 unit tests passing across 3 test files. Tool validation against live SharePoint and npm packaging pending (S03–S04).

## Architecture / Key Patterns

- **Runtime:** Node.js ESM, plain JavaScript (no TypeScript, no build step)
- **Protocol:** MCP over stdio (`@modelcontextprotocol/sdk`)
- **Auth:** MSAL Node (`@azure/msal-node`) with Device Code Flow
- **API:** Microsoft Graph v1.0 + Beta for pages/layout, SharePoint REST API for navigation/branding
- **Structure:**
  - `src/auth.js` — MSAL wrapper, token cache, device code flow, tenant discovery
  - `src/client.js` — Graph API + SP REST API client with per-resource token routing
  - `src/tools.js` — 25 MCP tool definitions (zod schemas + handlers), `parseSharePointUrl` export
  - `src/index.js` — Server entrypoint, stdio transport
  - `src/auth-cli.js` — Standalone auth test CLI

## Capability Contract

See `.gsd/REQUIREMENTS.md` for the explicit capability contract, requirement status, and coverage mapping.

## Milestone Sequence

- [ ] M001: Zero-Config SharePoint MCP — Refactor auth to zero-config, add site discovery, validate all tools, publish to npm as `sharepoint-online-mcp`
