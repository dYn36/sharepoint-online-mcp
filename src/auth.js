/**
 * SharePoint MCP - Authentication Module
 * Uses MSAL Device Code Flow (no admin rights needed)
 */

import { PublicClientApplication, CryptoProvider } from "@azure/msal-node";
import fs from "fs";
import path from "path";
import os from "os";

const TOKEN_CACHE_PATH = path.join(os.homedir(), ".sharepoint-mcp-cache.json");

// Default Azure AD App Registration scopes (delegated, no admin consent needed)
const DEFAULT_SCOPES = [
  "Sites.ReadWrite.All",
  "Sites.Manage.All",
  "User.Read",
];

/**
 * You need to register an app in Azure AD (portal.azure.com > App registrations).
 * - Set "Supported account types" to your org (single tenant) or multitenant
 * - Under "Authentication", enable "Allow public client flows" = Yes
 * - Under "API permissions", add Microsoft Graph delegated permissions:
 *   - Sites.ReadWrite.All
 *   - Sites.Manage.All  
 *   - User.Read
 * - NO admin consent needed for these delegated permissions
 */

export class SharePointAuth {
  constructor(clientId, tenantId = "common") {
    this.clientId = clientId;
    this.tenantId = tenantId;

    this.msalConfig = {
      auth: {
        clientId: this.clientId,
        authority: `https://login.microsoftonline.com/${this.tenantId}`,
      },
      cache: {
        cachePlugin: {
          beforeCacheAccess: async (context) => {
            if (fs.existsSync(TOKEN_CACHE_PATH)) {
              context.tokenCache.deserialize(
                fs.readFileSync(TOKEN_CACHE_PATH, "utf-8")
              );
            }
          },
          afterCacheAccess: async (context) => {
            if (context.cacheHasChanged) {
              fs.writeFileSync(TOKEN_CACHE_PATH, context.tokenCache.serialize());
            }
          },
        },
      },
    };

    this.pca = new PublicClientApplication(this.msalConfig);
    this.account = null;
  }

  /**
   * Try to get token silently from cache, fall back to device code flow
   */
  async getAccessToken(scopes = DEFAULT_SCOPES) {
    // Try silent acquisition first
    try {
      const accounts = await this.pca.getTokenCache().getAllAccounts();
      if (accounts.length > 0) {
        this.account = accounts[0];
        const result = await this.pca.acquireTokenSilent({
          account: this.account,
          scopes,
        });
        return result.accessToken;
      }
    } catch (e) {
      // Silent failed, proceed to device code
    }

    // Device Code Flow - works without admin rights
    const result = await this.pca.acquireTokenByDeviceCode({
      scopes,
      deviceCodeCallback: (response) => {
        // In MCP stdio mode, we write to stderr so it doesn't interfere with protocol
        process.stderr.write(
          `\n🔐 SharePoint Login erforderlich!\n` +
            `   Öffne: ${response.verificationUri}\n` +
            `   Code:  ${response.userCode}\n` +
            `   (Code läuft ab in ${Math.round(response.expiresIn / 60)} Minuten)\n\n`
        );
      },
    });

    this.account = result.account;
    return result.accessToken;
  }

  /**
   * Clear cached tokens
   */
  async logout() {
    if (fs.existsSync(TOKEN_CACHE_PATH)) {
      fs.unlinkSync(TOKEN_CACHE_PATH);
    }
    this.account = null;
  }
}
