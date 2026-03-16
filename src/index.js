#!/usr/bin/env node

/**
 * SharePoint Online MCP Server
 * Create, edit, and design SharePoint Online sites and pages via Claude.
 * Zero-config: authenticates via device code flow, no Azure Portal setup required.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { SharePointAuth } from "./auth.js";
import { SharePointClient } from "./client.js";
import { registerTools } from "./tools.js";

// ─── Initialize ───
const auth = new SharePointAuth();
const client = new SharePointClient(auth);

const server = new McpServer({
  name: "sharepoint-online-mcp",
  version: "1.0.0",
  description:
    "MCP Server for SharePoint Online — create, edit, and design SharePoint sites and pages via Claude. " +
    "Zero-config: authenticates via device code flow, no Azure Portal setup required.",
});

// Register all tools
registerTools(server, client, auth);

// ─── Start Server ───
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  process.stderr.write("🚀 SharePoint Online MCP Server started (stdio)\n");
}

main().catch((err) => {
  process.stderr.write(`❌ Startup error: ${err.message}\n`);
  process.exit(1);
});
