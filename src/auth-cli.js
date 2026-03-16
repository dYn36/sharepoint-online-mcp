#!/usr/bin/env node

/**
 * Standalone Auth Test — test authentication separately.
 * Usage: node src/auth-cli.js
 * No environment variables required (zero-config).
 */

import { SharePointAuth } from "./auth.js";

async function main() {
  console.log("🔐 Starting authentication...");
  console.log("   Using Microsoft Office well-known client ID (zero-config)\n");

  const auth = new SharePointAuth();

  try {
    const token = await auth.getAccessToken("https://graph.microsoft.com");
    console.log("\n✅ Authentication successful!");
    console.log(`   Token starts with: ${token.substring(0, 20)}...`);
    console.log(`   Account: ${auth.account?.username || "unknown"}`);
    console.log("\n   Token is cached. The MCP server won't prompt you again.");
  } catch (err) {
    console.error(`\n❌ Error: ${err.message}`);
  }
}

main();
