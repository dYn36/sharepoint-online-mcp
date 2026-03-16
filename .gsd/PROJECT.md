# Project

## What This Is

An MCP (Model Context Protocol) server that lets AI assistants design and edit SharePoint Online sites — pages, layouts, web parts, navigation, branding — using the authenticated user's own permissions. Zero-config: no Azure Portal, no app registration, no env vars.

## Core Value

A user runs `npx sharepoint-online-mcp` with zero configuration, authenticates via device code in their browser, and Claude can immediately create/edit/publish SharePoint pages on sites they have access to.

## Current State

Existing prototype with 20+ MCP tools covering pages, canvas layout, web parts, navigation, branding, and asset upload. All tools use Microsoft Graph API (beta) and SharePoint REST API. Auth currently requires manual Azure AD app registration and env vars (`SHAREPOINT_CLIENT_ID`, `SHAREPOINT_TENANT_ID`). Needs to be refactored to zero-config auth with well-known client IDs and automatic tenant discovery.

## Architecture / Key Patterns

- **Runtime:** Node.js ESM, plain JavaScript (no TypeScript, no build step)
- **Protocol:** MCP over stdio (`@modelcontextprotocol/sdk`)
- **Auth:** MSAL Node (`@azure/msal-node`) with Device Code Flow
- **API:** Microsoft Graph v1.0 + Beta for pages/layout, SharePoint REST API for navigation/branding
- **Structure:**
  - `src/auth.js` — MSAL wrapper, token cache, device code flow
  - `src/client.js` — Graph API + SP REST API client
  - `src/tools.js` — 20+ MCP tool definitions (zod schemas + handlers)
  - `src/index.js` — Server entrypoint, stdio transport
  - `src/auth-cli.js` — Standalone auth test CLI

## Capability Contract

See `.gsd/REQUIREMENTS.md` for the explicit capability contract, requirement status, and coverage mapping.

## Milestone Sequence

- [ ] M001: Zero-Config SharePoint MCP — Refactor auth to zero-config, add site discovery, validate all tools, publish to npm as `sharepoint-online-mcp`
