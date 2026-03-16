#!/usr/bin/env node

/**
 * Standalone Auth Test - teste die Authentifizierung separat
 * Usage: SHAREPOINT_CLIENT_ID=xxx SHAREPOINT_TENANT_ID=xxx node src/auth-cli.js
 */

import { SharePointAuth } from "./auth.js";

const CLIENT_ID = process.env.SHAREPOINT_CLIENT_ID;
const TENANT_ID = process.env.SHAREPOINT_TENANT_ID || "common";

if (!CLIENT_ID) {
  console.error("❌ SHAREPOINT_CLIENT_ID nicht gesetzt. Siehe README.md");
  process.exit(1);
}

async function main() {
  console.log("🔐 Starte Authentifizierung...");
  console.log(`   Client ID: ${CLIENT_ID}`);
  console.log(`   Tenant ID: ${TENANT_ID}\n`);

  const auth = new SharePointAuth(CLIENT_ID, TENANT_ID);

  try {
    const token = await auth.getAccessToken();
    console.log("\n✅ Authentifizierung erfolgreich!");
    console.log(`   Token beginnt mit: ${token.substring(0, 20)}...`);
    console.log(`   Account: ${auth.account?.username || "unbekannt"}`);
    console.log("\n   Token ist gecacht. Der MCP-Server wird dich nicht erneut fragen.");
  } catch (err) {
    console.error(`\n❌ Fehler: ${err.message}`);
  }
}

main();
