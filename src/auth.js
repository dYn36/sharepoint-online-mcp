/**
 * SharePoint MCP - Authentication Module
 * Zero-config auth engine using Microsoft Office well-known client ID.
 * Uses MSAL Device Code Flow (no admin rights needed).
 */

import { PublicClientApplication } from "@azure/msal-node";
import fs from "fs";
import path from "path";
import os from "os";

const TOKEN_CACHE_PATH = path.join(os.homedir(), ".sharepoint-mcp-cache.json");

/**
 * Microsoft Office well-known client ID.
 * Works with any Azure AD tenant — no app registration required.
 */
const WELL_KNOWN_CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c";

/**
 * Build OAuth2 scopes for a resource URL.
 * @param {string} resource - e.g. "https://graph.microsoft.com" or "https://contoso.sharepoint.com"
 * @returns {string[]} - e.g. ["https://graph.microsoft.com/.default"]
 */
export function buildScopes(resource) {
  const normalized = resource.endsWith("/")
    ? resource.slice(0, -1)
    : resource;
  return [`${normalized}/.default`];
}

/**
 * Discover the Azure AD tenant GUID for a SharePoint domain via OpenID configuration.
 * @param {string} domain - e.g. "contoso.sharepoint.com"
 * @param {Function} [fetchFn=fetch] - fetch implementation (injectable for testing)
 * @returns {Promise<string>} - tenant GUID
 */
export async function discoverTenantId(domain, fetchFn = fetch) {
  const url = `https://login.microsoftonline.com/${domain}/.well-known/openid-configuration`;

  let response;
  try {
    response = await fetchFn(url);
  } catch (err) {
    throw new Error(
      `Tenant discovery failed for "${domain}": network error — ${err.message}`
    );
  }

  if (!response.ok) {
    throw new Error(
      `Tenant discovery failed for "${domain}": HTTP ${response.status} from ${url}`
    );
  }

  let data;
  try {
    data = await response.json();
  } catch {
    throw new Error(
      `Tenant discovery failed for "${domain}": invalid JSON response from ${url}`
    );
  }

  const tokenEndpoint = data.token_endpoint;
  if (typeof tokenEndpoint !== "string") {
    throw new Error(
      `Tenant discovery failed for "${domain}": missing token_endpoint in OpenID config`
    );
  }

  // token_endpoint format: https://login.microsoftonline.com/{tenantGuid}/oauth2/v2.0/token
  const match = tokenEndpoint.match(
    /login\.microsoftonline\.com\/([0-9a-f-]{36})\//i
  );
  if (!match) {
    throw new Error(
      `Tenant discovery failed for "${domain}": could not parse tenant GUID from token_endpoint "${tokenEndpoint}"`
    );
  }

  return match[1];
}

/**
 * Zero-config SharePoint/Graph auth using device code flow.
 * No constructor arguments — uses well-known client ID and common authority.
 */
export class SharePointAuth {
  constructor() {
    this.msalConfig = {
      auth: {
        clientId: WELL_KNOWN_CLIENT_ID,
        authority: "https://login.microsoftonline.com/common",
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
              fs.writeFileSync(
                TOKEN_CACHE_PATH,
                context.tokenCache.serialize()
              );
            }
          },
        },
      },
    };

    this.pca = new PublicClientApplication(this.msalConfig);
    this.account = null;
  }

  /**
   * Acquire an access token for the given resource.
   * Tries silent cache acquisition first, falls back to device code flow.
   * @param {string} resource - e.g. "https://graph.microsoft.com" or "https://contoso.sharepoint.com"
   * @returns {Promise<string>} access token
   */
  async getAccessToken(resource) {
    const scopes = buildScopes(resource);

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
      // Silent acquisition failed — fall through to device code flow
      process.stderr.write(
        `[auth] Silent token acquisition failed for ${resource}: ${e.message}\n`
      );
    }

    // Device Code Flow - works without admin rights
    const result = await this.pca.acquireTokenByDeviceCode({
      scopes,
      deviceCodeCallback: (response) => {
        // In MCP stdio mode, write to stderr so it doesn't interfere with JSON-RPC protocol
        process.stderr.write(
          `\n🔐 SharePoint login required!\n` +
            `   Open:  ${response.verificationUri}\n` +
            `   Code:  ${response.userCode}\n` +
            `   (Code expires in ${Math.round(response.expiresIn / 60)} minutes)\n\n`
        );
      },
    });

    this.account = result.account;
    return result.accessToken;
  }

  /**
   * Clear cached tokens and reset auth state.
   */
  async logout() {
    if (fs.existsSync(TOKEN_CACHE_PATH)) {
      fs.unlinkSync(TOKEN_CACHE_PATH);
    }
    this.account = null;
  }
}
