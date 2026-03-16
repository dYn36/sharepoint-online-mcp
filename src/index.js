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
  process.stderr.write("\n");
  process.stderr.write("🚀 SharePoint Online MCP Server started (stdio)\n");
  process.stderr.write("\n");
  process.stderr.write("   This server communicates via MCP protocol over stdin/stdout.\n");
  process.stderr.write("   It is designed to be used with an MCP client like Claude Desktop.\n");
  process.stderr.write("\n");
  process.stderr.write("   📋 Claude Desktop config (~/.claude/claude_desktop_config.json):\n");
  process.stderr.write("\n");
  process.stderr.write('      { "mcpServers": { "sharepoint": { "command": "npx", "args": ["sharepoint-online-mcp"] } } }\n');
  process.stderr.write("\n");
  process.stderr.write("   Once connected, ask Claude to search for a SharePoint site to get started.\n");
  process.stderr.write("   Authentication happens automatically on first use (device code flow).\n");
  process.stderr.write("\n");
  process.stderr.write("   Waiting for MCP client connection...\n");
}

main().catch((err) => {
  process.stderr.write(`❌ Startup error: ${err.message}\n`);
  process.exit(1);
});
