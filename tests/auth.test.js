import { describe, it, mock, beforeEach, afterEach } from "node:test";
import assert from "node:assert/strict";

// We test the pure functions directly and the class construction.
// MSAL is not mocked — we only test the non-MSAL surface of auth.js.

import { discoverTenantId, buildScopes, SharePointAuth } from "../src/auth.js";

// ── buildScopes ──────────────────────────────────────────────────────────────

describe("buildScopes", () => {
  it("constructs .default scope for Graph resource", () => {
    assert.deepStrictEqual(
      buildScopes("https://graph.microsoft.com"),
      ["https://graph.microsoft.com/.default"]
    );
  });

  it("constructs .default scope for SharePoint resource", () => {
    assert.deepStrictEqual(
      buildScopes("https://contoso.sharepoint.com"),
      ["https://contoso.sharepoint.com/.default"]
    );
  });

  it("strips trailing slash before appending .default", () => {
    assert.deepStrictEqual(
      buildScopes("https://graph.microsoft.com/"),
      ["https://graph.microsoft.com/.default"]
    );
  });
});

// ── discoverTenantId ─────────────────────────────────────────────────────────

describe("discoverTenantId", () => {
  const TENANT_GUID = "72f988bf-86f1-41af-91ab-2d7cd011db47";
  const DOMAIN = "contoso.sharepoint.com";
  const EXPECTED_URL = `https://login.microsoftonline.com/${DOMAIN}/.well-known/openid-configuration`;

  it("fetches the correct OpenID configuration URL", async () => {
    const fakeFetch = mock.fn(async () => ({
      ok: true,
      json: async () => ({
        token_endpoint: `https://login.microsoftonline.com/${TENANT_GUID}/oauth2/v2.0/token`,
      }),
    }));

    await discoverTenantId(DOMAIN, fakeFetch);

    assert.strictEqual(fakeFetch.mock.calls.length, 1);
    assert.strictEqual(fakeFetch.mock.calls[0].arguments[0], EXPECTED_URL);
  });

  it("parses tenant GUID from token_endpoint", async () => {
    const fakeFetch = mock.fn(async () => ({
      ok: true,
      json: async () => ({
        token_endpoint: `https://login.microsoftonline.com/${TENANT_GUID}/oauth2/v2.0/token`,
      }),
    }));

    const result = await discoverTenantId(DOMAIN, fakeFetch);
    assert.strictEqual(result, TENANT_GUID);
  });

  it("throws with domain in message on network error", async () => {
    const fakeFetch = mock.fn(async () => {
      throw new Error("ENOTFOUND");
    });

    await assert.rejects(
      () => discoverTenantId(DOMAIN, fakeFetch),
      (err) => {
        assert.ok(err.message.includes(DOMAIN), `message should include domain: ${err.message}`);
        assert.ok(err.message.includes("network error"), `message should mention network error: ${err.message}`);
        return true;
      }
    );
  });

  it("throws with domain in message on HTTP error", async () => {
    const fakeFetch = mock.fn(async () => ({
      ok: false,
      status: 404,
    }));

    await assert.rejects(
      () => discoverTenantId(DOMAIN, fakeFetch),
      (err) => {
        assert.ok(err.message.includes(DOMAIN), `message should include domain: ${err.message}`);
        assert.ok(err.message.includes("404"), `message should include status code: ${err.message}`);
        return true;
      }
    );
  });

  it("throws on missing token_endpoint in response", async () => {
    const fakeFetch = mock.fn(async () => ({
      ok: true,
      json: async () => ({ issuer: "something" }),
    }));

    await assert.rejects(
      () => discoverTenantId(DOMAIN, fakeFetch),
      (err) => {
        assert.ok(err.message.includes(DOMAIN));
        assert.ok(err.message.includes("token_endpoint"));
        return true;
      }
    );
  });

  it("throws on unparseable tenant GUID in token_endpoint", async () => {
    const fakeFetch = mock.fn(async () => ({
      ok: true,
      json: async () => ({
        token_endpoint: "https://login.microsoftonline.com/not-a-guid/oauth2/v2.0/token",
      }),
    }));

    await assert.rejects(
      () => discoverTenantId(DOMAIN, fakeFetch),
      (err) => {
        assert.ok(err.message.includes(DOMAIN));
        assert.ok(err.message.includes("tenant GUID"));
        return true;
      }
    );
  });
});

// ── SharePointAuth construction ──────────────────────────────────────────────

describe("SharePointAuth", () => {
  it("constructs without any arguments", () => {
    const auth = new SharePointAuth();
    assert.ok(auth, "should instantiate");
  });

  it("exposes logout method", () => {
    const auth = new SharePointAuth();
    assert.strictEqual(typeof auth.logout, "function");
  });

  it("exposes getAccessToken method", () => {
    const auth = new SharePointAuth();
    assert.strictEqual(typeof auth.getAccessToken, "function");
  });
});
