/**
 * Unit tests for SharePointClient dual-audience token routing.
 * Verifies that request() passes Graph resource and spRest() passes SP origin to auth.
 */

import { describe, it, beforeEach, afterEach, mock } from "node:test";
import assert from "node:assert/strict";
import { SharePointClient } from "../src/client.js";

/**
 * Create a mock auth object that records getAccessToken calls.
 */
function createMockAuth() {
  const calls = [];
  return {
    calls,
    getAccessToken: async (resource) => {
      calls.push(resource);
      return "mock-token-for-" + resource;
    },
  };
}

/**
 * Create a mock fetch that returns a successful JSON response.
 */
function createMockFetch(body = {}) {
  return async () => ({
    ok: true,
    status: 200,
    text: async () => JSON.stringify(body),
  });
}

/**
 * Create a mock fetch that returns a configurable HTTP error response.
 */
function createErrorFetch(status, body = "") {
  return async () => ({
    ok: false,
    status,
    text: async () => body,
  });
}

describe("SharePointClient error messages (Graph API via request())", () => {
  it("401 → 'Authentication expired' guidance", async () => {
    const mockAuth = createMockAuth();
    const client = new SharePointClient(mockAuth, { fetchFn: createErrorFetch(401, "token expired") });

    await assert.rejects(
      () => client.graph("/me"),
      (err) => {
        assert.ok(err.message.includes("Authentication expired"), `expected 'Authentication expired', got: ${err.message}`);
        assert.ok(err.message.includes("disconnect"), `expected 'disconnect' guidance, got: ${err.message}`);
        return true;
      }
    );
  });

  it("403 → 'Access denied' guidance with original body", async () => {
    const mockAuth = createMockAuth();
    const client = new SharePointClient(mockAuth, { fetchFn: createErrorFetch(403, "insufficient privileges") });

    await assert.rejects(
      () => client.graph("/me"),
      (err) => {
        assert.ok(err.message.includes("Access denied"), `expected 'Access denied', got: ${err.message}`);
        assert.ok(err.message.includes("insufficient privileges"), `expected original body preserved, got: ${err.message}`);
        return true;
      }
    );
  });

  it("404 → 'Resource not found' guidance with original body", async () => {
    const mockAuth = createMockAuth();
    const client = new SharePointClient(mockAuth, { fetchFn: createErrorFetch(404, "item not found") });

    await assert.rejects(
      () => client.graph("/sites/nonexistent"),
      (err) => {
        assert.ok(err.message.includes("Resource not found"), `expected 'Resource not found', got: ${err.message}`);
        assert.ok(err.message.includes("item not found"), `expected original body preserved, got: ${err.message}`);
        return true;
      }
    );
  });

  it("500 → generic 'SharePoint API error' with status and body", async () => {
    const mockAuth = createMockAuth();
    const client = new SharePointClient(mockAuth, { fetchFn: createErrorFetch(500, "internal server error") });

    await assert.rejects(
      () => client.graph("/me"),
      (err) => {
        assert.ok(err.message.startsWith("SharePoint API error"), `expected 'SharePoint API error', got: ${err.message}`);
        assert.ok(err.message.includes("500"), `expected status code, got: ${err.message}`);
        assert.ok(err.message.includes("internal server error"), `expected body, got: ${err.message}`);
        return true;
      }
    );
  });
});

describe("SharePointClient error messages (SP REST API via spRest())", () => {
  it("401 → 'Authentication expired' guidance", async () => {
    const mockAuth = createMockAuth();
    const client = new SharePointClient(mockAuth, { fetchFn: createErrorFetch(401, "token expired") });

    await assert.rejects(
      () => client.spRest("https://contoso.sharepoint.com/sites/test", "web"),
      (err) => {
        assert.ok(err.message.includes("Authentication expired"), `expected 'Authentication expired', got: ${err.message}`);
        assert.ok(err.message.includes("disconnect"), `expected 'disconnect' guidance, got: ${err.message}`);
        return true;
      }
    );
  });

  it("403 → 'Access denied' guidance with original body", async () => {
    const mockAuth = createMockAuth();
    const client = new SharePointClient(mockAuth, { fetchFn: createErrorFetch(403, "access denied by policy") });

    await assert.rejects(
      () => client.spRest("https://contoso.sharepoint.com/sites/test", "web"),
      (err) => {
        assert.ok(err.message.includes("Access denied"), `expected 'Access denied', got: ${err.message}`);
        assert.ok(err.message.includes("access denied by policy"), `expected original body preserved, got: ${err.message}`);
        return true;
      }
    );
  });

  it("404 → 'Resource not found' guidance with original body", async () => {
    const mockAuth = createMockAuth();
    const client = new SharePointClient(mockAuth, { fetchFn: createErrorFetch(404, "list does not exist") });

    await assert.rejects(
      () => client.spRest("https://contoso.sharepoint.com/sites/test", "web/lists/getbytitle('Nonexistent')"),
      (err) => {
        assert.ok(err.message.includes("Resource not found"), `expected 'Resource not found', got: ${err.message}`);
        assert.ok(err.message.includes("list does not exist"), `expected original body preserved, got: ${err.message}`);
        return true;
      }
    );
  });

  it("500 → generic 'SharePoint API error' with status and body", async () => {
    const mockAuth = createMockAuth();
    const client = new SharePointClient(mockAuth, { fetchFn: createErrorFetch(500, "server error") });

    await assert.rejects(
      () => client.spRest("https://contoso.sharepoint.com/sites/test", "web"),
      (err) => {
        assert.ok(err.message.startsWith("SharePoint API error"), `expected 'SharePoint API error', got: ${err.message}`);
        assert.ok(err.message.includes("500"), `expected status code, got: ${err.message}`);
        assert.ok(err.message.includes("server error"), `expected body, got: ${err.message}`);
        return true;
      }
    );
  });
});

describe("SharePointClient token routing", () => {
  let originalFetch;

  beforeEach(() => {
    originalFetch = globalThis.fetch;
  });

  afterEach(() => {
    globalThis.fetch = originalFetch;
  });

  it("request() passes 'https://graph.microsoft.com' to getAccessToken", async () => {
    const mockAuth = createMockAuth();
    const client = new SharePointClient(mockAuth);

    // Mock the node-fetch import — client.js uses `import fetch from 'node-fetch'`
    // so we need to mock at the instance level. The client uses fetch directly,
    // but request() is what we test via graph(). We need to mock the module-level fetch.
    // Instead, we'll mock globalThis.fetch and verify that won't work since client.js
    // imports node-fetch. Let's test via the auth call tracking instead.

    // Since client.js imports node-fetch at module level, we can't easily replace it.
    // But we CAN verify the auth call by catching the fetch error — the token is acquired
    // before fetch is called. Let's make the test work by accepting the fetch will fail
    // and catching it.

    try {
      await client.graph("/me");
    } catch {
      // fetch will fail since node-fetch hits a real URL — that's fine,
      // we only care that getAccessToken was called with the right resource
    }

    assert.equal(mockAuth.calls.length, 1);
    assert.equal(mockAuth.calls[0], "https://graph.microsoft.com");
  });

  it("graphBetaReq() also passes 'https://graph.microsoft.com' to getAccessToken", async () => {
    const mockAuth = createMockAuth();
    const client = new SharePointClient(mockAuth);

    try {
      await client.graphBetaReq("/sites/root/pages");
    } catch {
      // fetch failure expected
    }

    assert.equal(mockAuth.calls.length, 1);
    assert.equal(mockAuth.calls[0], "https://graph.microsoft.com");
  });

  it("spRest() extracts origin from siteUrl and passes it to getAccessToken", async () => {
    const mockAuth = createMockAuth();
    const client = new SharePointClient(mockAuth);

    try {
      await client.spRest(
        "https://contoso.sharepoint.com/sites/marketing",
        "web"
      );
    } catch {
      // fetch failure expected
    }

    assert.equal(mockAuth.calls.length, 1);
    assert.equal(mockAuth.calls[0], "https://contoso.sharepoint.com");
  });

  it("spRest() handles -admin subdomain correctly", async () => {
    const mockAuth = createMockAuth();
    const client = new SharePointClient(mockAuth);

    try {
      await client.spRest(
        "https://contoso-admin.sharepoint.com/sites/admin-center",
        "web"
      );
    } catch {
      // fetch failure expected
    }

    assert.equal(mockAuth.calls.length, 1);
    assert.equal(mockAuth.calls[0], "https://contoso-admin.sharepoint.com");
  });

  it("spRest() handles -my subdomain (OneDrive) correctly", async () => {
    const mockAuth = createMockAuth();
    const client = new SharePointClient(mockAuth);

    try {
      await client.spRest(
        "https://contoso-my.sharepoint.com/personal/user",
        "web"
      );
    } catch {
      // fetch failure expected
    }

    assert.equal(mockAuth.calls.length, 1);
    assert.equal(mockAuth.calls[0], "https://contoso-my.sharepoint.com");
  });

  it("spRest() passes origin without trailing path segments", async () => {
    const mockAuth = createMockAuth();
    const client = new SharePointClient(mockAuth);

    try {
      await client.spRest(
        "https://fabrikam.sharepoint.com/sites/hr/subsite",
        "web/lists"
      );
    } catch {
      // fetch failure expected
    }

    assert.equal(mockAuth.calls.length, 1);
    assert.equal(mockAuth.calls[0], "https://fabrikam.sharepoint.com");
  });
});
