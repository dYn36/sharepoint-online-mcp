/**
 * SharePoint MCP - API Client
 * Wraps Microsoft Graph API + SharePoint REST API for site design
 */

import fetch from "node-fetch";

export class SharePointClient {
  constructor(auth) {
    this.auth = auth;
    this.graphBase = "https://graph.microsoft.com/v1.0";
    this.graphBeta = "https://graph.microsoft.com/beta";
  }

  async request(url, options = {}) {
    const token = await this.auth.getAccessToken();
    const res = await fetch(url, {
      ...options,
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        Accept: "application/json",
        ...options.headers,
      },
    });

    if (!res.ok) {
      const errBody = await res.text();
      throw new Error(`API ${res.status}: ${errBody}`);
    }

    const text = await res.text();
    return text ? JSON.parse(text) : null;
  }

  async graph(endpoint, options = {}) {
    return this.request(`${this.graphBase}${endpoint}`, options);
  }

  async graphBetaReq(endpoint, options = {}) {
    return this.request(`${this.graphBeta}${endpoint}`, options);
  }

  /**
   * SharePoint REST API call (for features not in Graph)
   */
  async spRest(siteUrl, apiPath, options = {}) {
    const token = await this.auth.getAccessToken();
    const url = `${siteUrl}/_api/${apiPath}`;
    const res = await fetch(url, {
      ...options,
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        Accept: "application/json;odata=verbose",
        ...options.headers,
      },
    });

    if (!res.ok) {
      const errBody = await res.text();
      throw new Error(`SP REST ${res.status}: ${errBody}`);
    }

    const text = await res.text();
    return text ? JSON.parse(text) : null;
  }

  // ─────────────────────────────────────────────
  // SITES
  // ─────────────────────────────────────────────

  async listSites(search = "") {
    if (search) {
      const data = await this.graph(`/sites?search=${encodeURIComponent(search)}`);
      return data.value;
    }
    // List sites the user follows or has access to
    const data = await this.graph("/me/followedSites");
    return data.value;
  }

  async getSite(siteId) {
    return this.graph(`/sites/${siteId}`);
  }

  async getSiteByUrl(hostname, sitePath) {
    return this.graph(`/sites/${hostname}:/${sitePath}`);
  }

  // ─────────────────────────────────────────────
  // PAGES (Site Pages via Graph Beta)
  // ─────────────────────────────────────────────

  async listPages(siteId) {
    const data = await this.graphBetaReq(`/sites/${siteId}/pages`);
    return data.value;
  }

  async getPage(siteId, pageId) {
    return this.graphBetaReq(
      `/sites/${siteId}/pages/${pageId}/microsoft.graph.sitePage?$expand=canvasLayout`
    );
  }

  async createPage(siteId, pageData) {
    return this.graphBetaReq(`/sites/${siteId}/pages`, {
      method: "POST",
      body: JSON.stringify({
        "@odata.type": "#microsoft.graph.sitePage",
        name: pageData.name,
        title: pageData.title,
        pageLayout: pageData.layout || "article",
        showComments: pageData.showComments ?? true,
        showRecommendedPages: pageData.showRecommendedPages ?? false,
        titleArea: pageData.titleArea || {
          enableGradientEffect: true,
          imageWebUrl: pageData.headerImageUrl || "",
          layout: pageData.titleLayout || "plain",
          showAuthor: true,
          showPublishedDate: true,
          showTextBlockAboveTitle: false,
          textAboveTitle: "",
          textAlignment: "left",
          title: pageData.title,
        },
        ...(pageData.canvasLayout ? { canvasLayout: pageData.canvasLayout } : {}),
      }),
    });
  }

  async updatePage(siteId, pageId, updates) {
    return this.graphBetaReq(`/sites/${siteId}/pages/${pageId}`, {
      method: "PATCH",
      body: JSON.stringify({
        "@odata.type": "#microsoft.graph.sitePage",
        ...updates,
      }),
    });
  }

  async publishPage(siteId, pageId) {
    return this.graphBetaReq(
      `/sites/${siteId}/pages/${pageId}/microsoft.graph.sitePage/publish`,
      { method: "POST" }
    );
  }

  async deletePage(siteId, pageId) {
    return this.graphBetaReq(`/sites/${siteId}/pages/${pageId}`, {
      method: "DELETE",
    });
  }

  // ─────────────────────────────────────────────
  // WEB PARTS (Page Content / Canvas)
  // ─────────────────────────────────────────────

  async getPageWebParts(siteId, pageId) {
    const data = await this.graphBetaReq(
      `/sites/${siteId}/pages/${pageId}/microsoft.graph.sitePage/webParts`
    );
    return data.value;
  }

  async getWebPart(siteId, pageId, webPartId) {
    return this.graphBetaReq(
      `/sites/${siteId}/pages/${pageId}/microsoft.graph.sitePage/webParts/${webPartId}`
    );
  }

  async createWebPartInSection(siteId, pageId, sectionInfo, webPartData) {
    const { sectionIndex, columnIndex } = sectionInfo;
    return this.graphBetaReq(
      `/sites/${siteId}/pages/${pageId}/microsoft.graph.sitePage/canvasLayout/horizontalSections/${sectionIndex}/columns/${columnIndex}/webparts`,
      {
        method: "POST",
        body: JSON.stringify(webPartData),
      }
    );
  }

  // ─────────────────────────────────────────────
  // CANVAS LAYOUT (Sections & Columns)
  // ─────────────────────────────────────────────

  async getCanvasLayout(siteId, pageId) {
    return this.graphBetaReq(
      `/sites/${siteId}/pages/${pageId}/microsoft.graph.sitePage/canvasLayout`
    );
  }

  async getHorizontalSections(siteId, pageId) {
    const data = await this.graphBetaReq(
      `/sites/${siteId}/pages/${pageId}/microsoft.graph.sitePage/canvasLayout/horizontalSections`
    );
    return data.value;
  }

  async addHorizontalSection(siteId, pageId, sectionData) {
    return this.graphBetaReq(
      `/sites/${siteId}/pages/${pageId}/microsoft.graph.sitePage/canvasLayout/horizontalSections`,
      {
        method: "POST",
        body: JSON.stringify(sectionData),
      }
    );
  }

  // ─────────────────────────────────────────────
  // SITE PROPERTIES & BRANDING
  // ─────────────────────────────────────────────

  async updateSiteProperties(siteId, properties) {
    return this.graph(`/sites/${siteId}`, {
      method: "PATCH",
      body: JSON.stringify(properties),
    });
  }

  async getSiteLists(siteId) {
    const data = await this.graph(`/sites/${siteId}/lists`);
    return data.value;
  }

  // ─────────────────────────────────────────────
  // NAVIGATION
  // ─────────────────────────────────────────────

  async getNavigation(siteUrl) {
    return this.spRest(siteUrl, "web/navigation/quicklaunch");
  }

  async getTopNavigation(siteUrl) {
    return this.spRest(siteUrl, "web/navigation/topnavigationbar");
  }

  async addNavigationNode(siteUrl, navType, nodeData) {
    const apiPath =
      navType === "top"
        ? "web/navigation/topnavigationbar"
        : "web/navigation/quicklaunch";
    return this.spRest(siteUrl, apiPath, {
      method: "POST",
      body: JSON.stringify({
        __metadata: { type: "SP.NavigationNode" },
        Title: nodeData.title,
        Url: nodeData.url,
        IsExternal: nodeData.isExternal || false,
      }),
    });
  }

  async deleteNavigationNode(siteUrl, navType, nodeId) {
    const apiPath =
      navType === "top"
        ? `web/navigation/topnavigationbar(${nodeId})`
        : `web/navigation/quicklaunch(${nodeId})`;
    return this.spRest(siteUrl, apiPath, { method: "DELETE" });
  }

  // ─────────────────────────────────────────────
  // SITE COLUMNS & CONTENT TYPES
  // ─────────────────────────────────────────────

  async getSiteColumns(siteId) {
    const data = await this.graph(`/sites/${siteId}/columns`);
    return data.value;
  }

  async getContentTypes(siteId) {
    const data = await this.graph(`/sites/${siteId}/contentTypes`);
    return data.value;
  }

  // ─────────────────────────────────────────────
  // SITE THEMING (via SharePoint REST)
  // ─────────────────────────────────────────────

  async getSiteTheming(siteUrl) {
    return this.spRest(siteUrl, "web?$select=ThemedCssFolderUrl,ThemeInfo");
  }

  async setSiteLogo(siteUrl, logoUrl) {
    return this.spRest(siteUrl, "web", {
      method: "POST",
      headers: {
        "X-HTTP-Method": "MERGE",
      },
      body: JSON.stringify({
        __metadata: { type: "SP.Web" },
        SiteLogoUrl: logoUrl,
      }),
    });
  }

  // ─────────────────────────────────────────────
  // FILES (for uploading images/assets)
  // ─────────────────────────────────────────────

  async uploadFile(siteId, folderPath, fileName, content) {
    const driveItem = await this.graph(
      `/sites/${siteId}/drive/root:/${folderPath}/${fileName}:/content`,
      {
        method: "PUT",
        headers: { "Content-Type": "application/octet-stream" },
        body: content,
      }
    );
    return driveItem;
  }

  async getFileUrl(siteId, filePath) {
    const item = await this.graph(
      `/sites/${siteId}/drive/root:/${filePath}`
    );
    return item;
  }
}
