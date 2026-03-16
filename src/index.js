#!/usr/bin/env node

/**
 * SharePoint MCP Server
 * Ermöglicht das Bearbeiten und Designen von SharePoint Online Sites
 * über das Model Context Protocol (stdio).
 *
 * Benötigt:
 *   SHAREPOINT_CLIENT_ID  - Azure AD App Registration Client ID
 *   SHAREPOINT_TENANT_ID  - Azure AD Tenant ID (optional, default: "common")
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { SharePointAuth } from "./auth.js";
import { SharePointClient } from "./client.js";
import { registerTools } from "./tools.js";

// ─── Config ───
const CLIENT_ID = process.env.SHAREPOINT_CLIENT_ID;
const TENANT_ID = process.env.SHAREPOINT_TENANT_ID || "common";

if (!CLIENT_ID) {
  process.stderr.write(
    `\n❌ Fehler: SHAREPOINT_CLIENT_ID ist nicht gesetzt!\n\n` +
      `So richtest du es ein:\n` +
      `1. Gehe zu https://portal.azure.com → Azure Active Directory → App registrations\n` +
      `2. Klicke "New registration"\n` +
      `3. Name: "SharePoint MCP" (oder beliebig)\n` +
      `4. Supported account types: "Accounts in this organizational directory only"\n` +
      `5. Redirect URI: leer lassen\n` +
      `6. Nach Erstellung: Kopiere die "Application (client) ID"\n` +
      `7. Unter "Authentication": Setze "Allow public client flows" auf "Yes"\n` +
      `8. Unter "API permissions" → "Add a permission" → "Microsoft Graph" → "Delegated":\n` +
      `   - Sites.ReadWrite.All\n` +
      `   - Sites.Manage.All\n` +
      `   - User.Read\n` +
      `9. Setze die Umgebungsvariablen:\n` +
      `   export SHAREPOINT_CLIENT_ID="deine-client-id"\n` +
      `   export SHAREPOINT_TENANT_ID="deine-tenant-id"\n\n`
  );
  process.exit(1);
}

// ─── Initialize ───
const auth = new SharePointAuth(CLIENT_ID, TENANT_ID);
const client = new SharePointClient(auth);

const server = new McpServer({
  name: "sharepoint-mcp",
  version: "1.0.0",
  description:
    "MCP Server für SharePoint Online – Sites bearbeiten und designen (UI/UX). " +
    "Erstellt, bearbeitet und published SharePoint-Seiten mit Sections, Web Parts, " +
    "Navigation und Branding. Nutzt Microsoft Graph API mit delegierten Berechtigungen.",
});

// Register all tools
registerTools(server, client);

// ─── Start Server ───
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  process.stderr.write("🚀 SharePoint MCP Server gestartet (stdio)\n");
}

main().catch((err) => {
  process.stderr.write(`❌ Startfehler: ${err.message}\n`);
  process.exit(1);
});
