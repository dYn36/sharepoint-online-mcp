import { describe, it, beforeEach } from "node:test";
import assert from "node:assert/strict";
import { parseSharePointUrl, registerTools } from "../src/tools.js";

// ─── Mock Infrastructure ───

class MockServer {
  constructor() {
    this.tools = new Map();
  }
  tool(name, _desc, _schema, handler) {
    this.tools.set(name, handler);
  }
  async call(name, args) {
    const handler = this.tools.get(name);
    if (!handler) throw new Error(`Tool not registered: ${name}`);
    return handler(args);
  }
}

class MockClient {
  constructor() {
    this.calls = [];
    this._listSitesError = null;
  }

  _record(method, ...args) {
    this.calls.push([method, ...args]);
  }

  async listSites(query) {
    this._record("listSites", query);
    if (this._listSitesError) throw this._listSitesError;
    return [{ id: "s1", displayName: "Site", webUrl: "https://x", description: "d" }];
  }
  async getSite(siteId) {
    this._record("getSite", siteId);
    return { id: siteId, displayName: "Site", webUrl: "https://x", description: "d", createdDateTime: "2024-01-01" };
  }
  async getSiteLists(siteId) {
    this._record("getSiteLists", siteId);
    return [{ id: "l1", displayName: "Documents", list: { template: "documentLibrary" } }];
  }
  async getSiteByUrl(hostname, path) {
    this._record("getSiteByUrl", hostname, path);
    return { id: "s1", displayName: "Site", webUrl: "https://x", description: "d" };
  }
  async listPages(siteId) {
    this._record("listPages", siteId);
    return [{ id: "p1", name: "Home.aspx", title: "Home", webUrl: "https://x/p", pageLayout: "article", lastModifiedDateTime: "2024-01-01" }];
  }
  async getPage(siteId, pageId) {
    this._record("getPage", siteId, pageId);
    return { id: pageId, title: "Page", webUrl: "https://x" };
  }
  async createPage(siteId, opts) {
    this._record("createPage", siteId, opts);
    return { id: "newpage", name: opts.name, title: opts.title, webUrl: "https://x/p" };
  }
  async updatePage(siteId, pageId, updates) {
    this._record("updatePage", siteId, pageId, updates);
  }
  async publishPage(siteId, pageId) {
    this._record("publishPage", siteId, pageId);
  }
  async deletePage(siteId, pageId) {
    this._record("deletePage", siteId, pageId);
  }
  async addHorizontalSection(siteId, pageId, template) {
    this._record("addHorizontalSection", siteId, pageId, template);
    return { id: "sec1" };
  }
  async getHorizontalSections(siteId, pageId) {
    this._record("getHorizontalSections", siteId, pageId);
    return [{ id: "sec1", layout: "fullWidth" }];
  }
  async getPageWebParts(siteId, pageId) {
    this._record("getPageWebParts", siteId, pageId);
    return [{ id: "wp1", type: "text" }];
  }
  async createWebPartInSection(siteId, pageId, pos, webPart) {
    this._record("createWebPartInSection", siteId, pageId, pos, webPart);
    return { id: "wp-new" };
  }
  async getNavigation(siteUrl) {
    this._record("getNavigation", siteUrl);
    return [{ Id: 1, Title: "Home", Url: "/" }];
  }
  async getTopNavigation(siteUrl) {
    this._record("getTopNavigation", siteUrl);
    return [{ Id: 2, Title: "About", Url: "/about" }];
  }
  async addNavigationNode(siteUrl, navType, node) {
    this._record("addNavigationNode", siteUrl, navType, node);
  }
  async deleteNavigationNode(siteUrl, navType, nodeId) {
    this._record("deleteNavigationNode", siteUrl, navType, nodeId);
  }
  async setSiteLogo(siteUrl, logoUrl) {
    this._record("setSiteLogo", siteUrl, logoUrl);
  }
  async uploadFile(siteId, folder, fileName, buffer) {
    this._record("uploadFile", siteId, folder, fileName, buffer);
    return { id: "f1", webUrl: "https://x/f" };
  }
}

class MockAuth {
  constructor() {
    this.logoutCalled = false;
  }
  async logout() {
    this.logoutCalled = true;
  }
}

const mockDiscoverTenantId = async (_hostname) => "fake-tenant-id-123";

// ─── Helpers ───

function assertMcpContent(result) {
  assert.ok(result.content, "result should have content");
  assert.ok(Array.isArray(result.content), "content should be array");
  assert.ok(result.content.length > 0, "content should not be empty");
  assert.equal(result.content[0].type, "text");
  assert.equal(typeof result.content[0].text, "string");
}

function parseContent(result) {
  return JSON.parse(result.content[0].text);
}

function findCalls(client, method) {
  return client.calls.filter((c) => c[0] === method);
}

// ─── Register tools once for all handler tests ───

const server = new MockServer();
const client = new MockClient();
const auth = new MockAuth();
registerTools(server, client, auth, { discoverTenantId: mockDiscoverTenantId });

// ─── URL Parsing (existing tests) ───

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

// ─── Tool Handler Tests ───

describe("Tool handlers", () => {
  beforeEach(() => {
    client.calls = [];
    client._listSitesError = null;
    auth.logoutCalled = false;
  });

  // ═══ SITES ═══

  describe("Sites", () => {
    it("search_sites calls listSites and returns mapped array", async () => {
      const result = await server.call("search_sites", { query: "marketing" });
      assertMcpContent(result);
      const data = parseContent(result);
      assert.ok(Array.isArray(data));
      assert.equal(data[0].id, "s1");
      assert.equal(data[0].name, "Site");
      assert.equal(data[0].url, "https://x");
      assert.deepStrictEqual(findCalls(client, "listSites")[0], ["listSites", "marketing"]);
    });

    it("get_site_details calls getSite and getSiteLists in parallel", async () => {
      const result = await server.call("get_site_details", { siteId: "site-abc" });
      assertMcpContent(result);
      const data = parseContent(result);
      assert.equal(data.site.id, "site-abc");
      assert.equal(data.site.name, "Site");
      assert.ok(Array.isArray(data.lists));
      assert.equal(data.lists[0].id, "l1");
      assert.equal(data.lists[0].template, "documentLibrary");
      assert.ok(findCalls(client, "getSite").length === 1);
      assert.ok(findCalls(client, "getSiteLists").length === 1);
      assert.deepStrictEqual(findCalls(client, "getSite")[0], ["getSite", "site-abc"]);
      assert.deepStrictEqual(findCalls(client, "getSiteLists")[0], ["getSiteLists", "site-abc"]);
    });

    it("get_site_by_url calls getSiteByUrl with hostname and path", async () => {
      const result = await server.call("get_site_by_url", { hostname: "contoso.sharepoint.com", sitePath: "sites/hr" });
      assertMcpContent(result);
      const data = parseContent(result);
      assert.equal(data.id, "s1");
      assert.deepStrictEqual(findCalls(client, "getSiteByUrl")[0], ["getSiteByUrl", "contoso.sharepoint.com", "sites/hr"]);
    });

    it("connect_to_site calls discoverTenantId then getSiteByUrl", async () => {
      const result = await server.call("connect_to_site", { url: "https://contoso.sharepoint.com/sites/test" });
      assertMcpContent(result);
      const data = parseContent(result);
      assert.equal(data.tenantId, "fake-tenant-id-123");
      assert.equal(data.id, "s1");
      assert.ok(findCalls(client, "getSiteByUrl").length === 1);
    });

    it("list_my_sites calls listSites with empty string", async () => {
      const result = await server.call("list_my_sites", {});
      assertMcpContent(result);
      const data = parseContent(result);
      assert.ok(Array.isArray(data));
      assert.equal(data[0].id, "s1");
      assert.deepStrictEqual(findCalls(client, "listSites")[0], ["listSites", ""]);
    });
  });

  // ═══ PAGES ═══

  describe("Pages", () => {
    it("list_pages calls listPages and returns mapped page array", async () => {
      const result = await server.call("list_pages", { siteId: "site-1" });
      assertMcpContent(result);
      const data = parseContent(result);
      assert.ok(Array.isArray(data));
      assert.equal(data[0].id, "p1");
      assert.equal(data[0].name, "Home.aspx");
      assert.equal(data[0].pageLayout, "article");
      assert.deepStrictEqual(findCalls(client, "listPages")[0], ["listPages", "site-1"]);
    });

    it("get_page calls getPage with siteId and pageId", async () => {
      const result = await server.call("get_page", { siteId: "site-1", pageId: "page-1" });
      assertMcpContent(result);
      const data = parseContent(result);
      assert.equal(data.id, "page-1");
      assert.deepStrictEqual(findCalls(client, "getPage")[0], ["getPage", "site-1", "page-1"]);
    });

    it("create_page calls createPage with name and title", async () => {
      const result = await server.call("create_page", {
        siteId: "site-1", name: "about.aspx", title: "About Us",
      });
      assertMcpContent(result);
      const data = parseContent(result);
      assert.equal(data.name, "about.aspx");
      assert.equal(data.title, "About Us");
      assert.equal(data.status, "draft");
      const call = findCalls(client, "createPage")[0];
      assert.equal(call[1], "site-1");
      assert.equal(call[2].name, "about.aspx");
      assert.equal(call[2].title, "About Us");
      assert.equal(findCalls(client, "publishPage").length, 0);
    });

    it("create_page with autoPublish calls createPage then publishPage", async () => {
      const result = await server.call("create_page", {
        siteId: "site-1", name: "news.aspx", title: "News", autoPublish: true,
      });
      assertMcpContent(result);
      const data = parseContent(result);
      assert.equal(data.status, "published");
      assert.equal(findCalls(client, "createPage").length, 1);
      assert.equal(findCalls(client, "publishPage").length, 1);
      assert.deepStrictEqual(findCalls(client, "publishPage")[0], ["publishPage", "site-1", "newpage"]);
    });

    it("delete_page calls deletePage with siteId and pageId", async () => {
      const result = await server.call("delete_page", { siteId: "site-1", pageId: "page-99" });
      assertMcpContent(result);
      assert.ok(result.content[0].text.includes("page-99"));
      assert.deepStrictEqual(findCalls(client, "deletePage")[0], ["deletePage", "site-1", "page-99"]);
    });
  });

  // ═══ LAYOUT ═══

  describe("Layout", () => {
    it("add_section calls addHorizontalSection with template", async () => {
      const result = await server.call("add_section", {
        siteId: "site-1", pageId: "page-1", sectionTemplate: "fullWidth",
      });
      assertMcpContent(result);
      const call = findCalls(client, "addHorizontalSection")[0];
      assert.equal(call[1], "site-1");
      assert.equal(call[2], "page-1");
      assert.equal(call[3].layout, "fullWidth");
      assert.equal(call[3].emphasis, "none");
    });

    it("add_section with emphasis overrides default", async () => {
      await server.call("add_section", {
        siteId: "site-1", pageId: "page-1", sectionTemplate: "twoColumns", emphasis: "strong",
      });
      const call = findCalls(client, "addHorizontalSection")[0];
      assert.equal(call[3].layout, "twoColumn");
      assert.equal(call[3].emphasis, "strong");
    });

    it("get_page_layout calls getHorizontalSections and getPageWebParts", async () => {
      const result = await server.call("get_page_layout", { siteId: "site-1", pageId: "page-1" });
      assertMcpContent(result);
      const data = parseContent(result);
      assert.ok(Array.isArray(data.sections));
      assert.ok(Array.isArray(data.webParts));
      assert.equal(data.sections[0].id, "sec1");
      assert.equal(data.webParts[0].id, "wp1");
      assert.equal(findCalls(client, "getHorizontalSections").length, 1);
      assert.equal(findCalls(client, "getPageWebParts").length, 1);
    });
  });

  // ═══ WEB PARTS ═══

  describe("Web Parts", () => {
    it("add_text_webpart calls createWebPartInSection with text template", async () => {
      const result = await server.call("add_text_webpart", {
        siteId: "site-1", pageId: "page-1", sectionIndex: 1, columnIndex: 1, html: "<p>Hello</p>",
      });
      assertMcpContent(result);
      const call = findCalls(client, "createWebPartInSection")[0];
      assert.equal(call[1], "site-1");
      assert.equal(call[2], "page-1");
      assert.deepStrictEqual(call[3], { sectionIndex: 1, columnIndex: 1 });
      assert.equal(call[4]["@odata.type"], "#microsoft.graph.textWebPart");
      assert.equal(call[4].innerHtml, "<p>Hello</p>");
    });

    it("add_image_webpart calls createWebPartInSection with image template", async () => {
      const result = await server.call("add_image_webpart", {
        siteId: "site-1", pageId: "page-1", sectionIndex: 2, columnIndex: 1,
        imageUrl: "https://img.example.com/hero.jpg", altText: "Hero", caption: "Banner",
      });
      assertMcpContent(result);
      const call = findCalls(client, "createWebPartInSection")[0];
      assert.deepStrictEqual(call[3], { sectionIndex: 2, columnIndex: 1 });
      assert.equal(call[4]["@odata.type"], "#microsoft.graph.standardWebPart");
      assert.equal(call[4].data.properties.imgUrl, "https://img.example.com/hero.jpg");
      assert.equal(call[4].data.properties.altText, "Hero");
      assert.equal(call[4].data.properties.caption, "Banner");
    });

    it("add_spacer calls createWebPartInSection with spacer template", async () => {
      const result = await server.call("add_spacer", {
        siteId: "site-1", pageId: "page-1", sectionIndex: 1, columnIndex: 1, height: 80,
      });
      assertMcpContent(result);
      const call = findCalls(client, "createWebPartInSection")[0];
      assert.equal(call[4].data.properties.height, 80);
    });

    it("add_divider calls createWebPartInSection with divider template", async () => {
      const result = await server.call("add_divider", {
        siteId: "site-1", pageId: "page-1", sectionIndex: 1, columnIndex: 1,
      });
      assertMcpContent(result);
      const call = findCalls(client, "createWebPartInSection")[0];
      assert.equal(call[4]["@odata.type"], "#microsoft.graph.standardWebPart");
      // Divider has the specific container ID
      assert.equal(call[4].containerTextWebPartId, "2161a1c6-db61-4731-b97c-3cdb303f7cbb");
    });

    it("add_custom_webpart calls createWebPartInSection with parsed JSON", async () => {
      const custom = { "@odata.type": "#microsoft.graph.standardWebPart", data: { custom: true } };
      const result = await server.call("add_custom_webpart", {
        siteId: "site-1", pageId: "page-1", sectionIndex: 3, columnIndex: 2,
        webPartJson: JSON.stringify(custom),
      });
      assertMcpContent(result);
      const call = findCalls(client, "createWebPartInSection")[0];
      assert.deepStrictEqual(call[3], { sectionIndex: 3, columnIndex: 2 });
      assert.deepStrictEqual(call[4], custom);
    });

    it("web part position has sectionIndex and columnIndex shape", async () => {
      await server.call("add_text_webpart", {
        siteId: "site-1", pageId: "page-1", sectionIndex: 5, columnIndex: 3, html: "<p>x</p>",
      });
      const call = findCalls(client, "createWebPartInSection")[0];
      assert.ok("sectionIndex" in call[3]);
      assert.ok("columnIndex" in call[3]);
      assert.equal(call[3].sectionIndex, 5);
      assert.equal(call[3].columnIndex, 3);
    });
  });

  // ═══ NAVIGATION ═══

  describe("Navigation", () => {
    it("get_navigation with quick calls getNavigation (not getTopNavigation)", async () => {
      const result = await server.call("get_navigation", {
        siteUrl: "https://contoso.sharepoint.com/sites/hr", navType: "quick",
      });
      assertMcpContent(result);
      const data = parseContent(result);
      assert.equal(data[0].Title, "Home");
      assert.equal(findCalls(client, "getNavigation").length, 1);
      assert.equal(findCalls(client, "getTopNavigation").length, 0);
    });

    it("get_navigation with top calls getTopNavigation", async () => {
      const result = await server.call("get_navigation", {
        siteUrl: "https://contoso.sharepoint.com/sites/hr", navType: "top",
      });
      assertMcpContent(result);
      const data = parseContent(result);
      assert.equal(data[0].Title, "About");
      assert.equal(findCalls(client, "getTopNavigation").length, 1);
      assert.equal(findCalls(client, "getNavigation").length, 0);
    });

    it("add_navigation_link calls addNavigationNode with navType and node", async () => {
      const result = await server.call("add_navigation_link", {
        siteUrl: "https://contoso.sharepoint.com/sites/hr", navType: "quick",
        title: "FAQ", url: "/faq", isExternal: false,
      });
      assertMcpContent(result);
      assert.ok(result.content[0].text.includes("FAQ"));
      const call = findCalls(client, "addNavigationNode")[0];
      assert.equal(call[1], "https://contoso.sharepoint.com/sites/hr");
      assert.equal(call[2], "quick");
      assert.deepStrictEqual(call[3], { title: "FAQ", url: "/faq", isExternal: false });
    });

    it("remove_navigation_link calls deleteNavigationNode", async () => {
      const result = await server.call("remove_navigation_link", {
        siteUrl: "https://contoso.sharepoint.com/sites/hr", navType: "top", nodeId: 42,
      });
      assertMcpContent(result);
      assert.ok(result.content[0].text.includes("42"));
      assert.deepStrictEqual(findCalls(client, "deleteNavigationNode")[0], [
        "deleteNavigationNode", "https://contoso.sharepoint.com/sites/hr", "top", 42,
      ]);
    });
  });

  // ═══ BRANDING ═══

  describe("Branding", () => {
    it("set_site_logo calls setSiteLogo with siteUrl and logoUrl", async () => {
      const result = await server.call("set_site_logo", {
        siteUrl: "https://contoso.sharepoint.com/sites/hr",
        logoUrl: "https://cdn.example.com/logo.png",
      });
      assertMcpContent(result);
      assert.deepStrictEqual(findCalls(client, "setSiteLogo")[0], [
        "setSiteLogo", "https://contoso.sharepoint.com/sites/hr", "https://cdn.example.com/logo.png",
      ]);
    });

    it("upload_asset calls uploadFile with decoded base64 buffer", async () => {
      const b64 = Buffer.from("fake-image-data").toString("base64");
      const result = await server.call("upload_asset", {
        siteId: "site-1", fileName: "hero.jpg", base64Content: b64,
      });
      assertMcpContent(result);
      const data = parseContent(result);
      assert.equal(data.webUrl, "https://x/f");
      const call = findCalls(client, "uploadFile")[0];
      assert.equal(call[1], "site-1");
      assert.equal(call[2], "SiteAssets"); // default folder
      assert.equal(call[3], "hero.jpg");
      assert.ok(Buffer.isBuffer(call[4]));
      assert.equal(call[4].toString(), "fake-image-data");
    });
  });

  // ═══ AUTH & UTILITY ═══

  describe("Auth & Utility", () => {
    it("disconnect calls auth.logout", async () => {
      const result = await server.call("disconnect", {});
      assertMcpContent(result);
      assert.ok(result.content[0].text.includes("Disconnected"));
      assert.ok(auth.logoutCalled);
    });

    it("get_design_templates returns static data with sectionLayouts", async () => {
      const result = await server.call("get_design_templates", {});
      assertMcpContent(result);
      const data = parseContent(result);
      assert.ok(Array.isArray(data.sectionLayouts));
      assert.ok(data.sectionLayouts.includes("fullWidth"));
      assert.ok(data.sectionLayouts.includes("threeColumns"));
      assert.ok(Array.isArray(data.titleAreaLayouts));
      assert.ok(Array.isArray(data.webPartTypes));
      assert.ok(Array.isArray(data.tips));
      // No client calls for static data
      assert.equal(client.calls.length, 0);
    });

    it("connect_to_site returns isError on invalid URL", async () => {
      const result = await server.call("connect_to_site", { url: "not-a-valid-url" });
      assert.ok(result.isError);
      assertMcpContent(result);
      assert.ok(result.content[0].text.includes("Error"));
    });

    it("list_my_sites returns isError when client throws", async () => {
      client._listSitesError = new Error("Network failure");
      const result = await server.call("list_my_sites", {});
      assert.ok(result.isError);
      assertMcpContent(result);
      assert.ok(result.content[0].text.includes("Network failure"));
    });
  });

  // ═══ ADDITIONAL PAGE TOOLS ═══

  describe("Additional page tools", () => {
    it("update_page calls updatePage with updates object", async () => {
      const result = await server.call("update_page", {
        siteId: "site-1", pageId: "page-1", title: "New Title",
      });
      assertMcpContent(result);
      const call = findCalls(client, "updatePage")[0];
      assert.equal(call[1], "site-1");
      assert.equal(call[2], "page-1");
      assert.equal(call[3].title, "New Title");
    });

    it("publish_page calls publishPage", async () => {
      const result = await server.call("publish_page", { siteId: "site-1", pageId: "page-2" });
      assertMcpContent(result);
      assert.ok(result.content[0].text.includes("page-2"));
      assert.deepStrictEqual(findCalls(client, "publishPage")[0], ["publishPage", "site-1", "page-2"]);
    });
  });
});
