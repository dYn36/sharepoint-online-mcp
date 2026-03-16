import { describe, it } from "node:test";
import assert from "node:assert/strict";
import { parseSharePointUrl } from "../src/tools.js";

describe("parseSharePointUrl", () => {
  it("parses standard site URL", () => {
    const result = parseSharePointUrl("https://contoso.sharepoint.com/sites/marketing");
    assert.deepStrictEqual(result, { hostname: "contoso.sharepoint.com", sitePath: "sites/marketing" });
  });

  it("parses root site URL", () => {
    const result = parseSharePointUrl("https://contoso.sharepoint.com");
    assert.deepStrictEqual(result, { hostname: "contoso.sharepoint.com", sitePath: "" });
  });

  it("strips trailing slash", () => {
    const result = parseSharePointUrl("https://contoso.sharepoint.com/sites/marketing/");
    assert.deepStrictEqual(result, { hostname: "contoso.sharepoint.com", sitePath: "sites/marketing" });
  });

  it("ignores subpage path beyond site path", () => {
    const result = parseSharePointUrl("https://contoso.sharepoint.com/sites/marketing/SitePages/Home.aspx");
    assert.deepStrictEqual(result, { hostname: "contoso.sharepoint.com", sitePath: "sites/marketing" });
  });

  it("parses teams prefix", () => {
    const result = parseSharePointUrl("https://contoso.sharepoint.com/teams/engineering");
    assert.deepStrictEqual(result, { hostname: "contoso.sharepoint.com", sitePath: "teams/engineering" });
  });

  it("throws on invalid URL", () => {
    assert.throws(() => parseSharePointUrl("not-a-url"), /Invalid SharePoint URL/);
  });

  it("parses URL with port", () => {
    const result = parseSharePointUrl("https://contoso.sharepoint.com:443/sites/test");
    assert.deepStrictEqual(result, { hostname: "contoso.sharepoint.com", sitePath: "sites/test" });
  });
});
