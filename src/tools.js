/**
 * SharePoint MCP - Tool Definitions & Handlers
 */

import { z } from "zod";

// ─── Common Web Part Templates ───
const WEB_PART_TEMPLATES = {
  text: (html) => ({
    "@odata.type": "#microsoft.graph.textWebPart",
    innerHtml: html,
  }),
  image: (imageUrl, altText = "", caption = "") => ({
    "@odata.type": "#microsoft.graph.standardWebPart",
    containerTextWebPartId: "d1d91016-032f-456d-98a4-721247c305e8",
    data: {
      dataVersion: "1.9",
      properties: {
        imageSourceType: 2,
        imgUrl: imageUrl,
        altText,
        caption,
      },
    },
  }),
  hero: (items) => ({
    "@odata.type": "#microsoft.graph.standardWebPart",
    containerTextWebPartId: "c4bd7b2f-7b6e-4599-8b44-4b01e7276f3b",
    data: {
      dataVersion: "1.0",
      properties: {
        heroLayoutThreshold: 640,
        carouselLayoutMaxWidth: 639,
        layoutCategory: 1,
        layout: 5,
        content: { items },
      },
    },
  }),
  quickLinks: (items) => ({
    "@odata.type": "#microsoft.graph.standardWebPart",
    containerTextWebPartId: "c70391ea-0b10-4ee9-b2b4-006d3fcad0cd",
    data: {
      dataVersion: "2.2",
      properties: {
        items,
        layoutId: "CompactCard",
        shouldShowThumbnail: true,
        hideWebPartWhenEmpty: true,
        dataProviderId: "QuickLinks",
      },
    },
  }),
  spacer: (height = 60) => ({
    "@odata.type": "#microsoft.graph.standardWebPart",
    containerTextWebPartId: "8f7d38f1-32cb-40f5-b7ad-c4382ac5a999",
    data: {
      dataVersion: "1.0",
      properties: { height },
    },
  }),
  divider: () => ({
    "@odata.type": "#microsoft.graph.standardWebPart",
    containerTextWebPartId: "2161a1c6-db61-4731-b97c-3cdb303f7cbb",
    data: { dataVersion: "1.0", properties: {} },
  }),
};

// ─── Section Layout Templates ───
const SECTION_TEMPLATES = {
  fullWidth: { layout: "fullWidth", emphasis: "none" },
  oneColumn: { layout: "oneColumn", emphasis: "none" },
  twoColumns: { layout: "twoColumn", emphasis: "none" },
  twoColumnsLeft: { layout: "twoColumnLeft", emphasis: "none" },
  twoColumnsRight: { layout: "twoColumnRight", emphasis: "none" },
  threeColumns: { layout: "threeColumn", emphasis: "none" },
  oneThirdLeft: { layout: "oneThirdLeft", emphasis: "none" },
  oneThirdRight: { layout: "oneThirdRight", emphasis: "none" },
};

// ─── Title Area Templates ───
const TITLE_TEMPLATES = {
  plain: {
    layout: "plain",
    showAuthor: true,
    showPublishedDate: true,
    textAlignment: "left",
  },
  imageAndTitle: {
    layout: "imageAndTitle",
    enableGradientEffect: true,
    showAuthor: true,
    showPublishedDate: true,
    textAlignment: "left",
    imageWebUrl: "",
  },
  colorBlock: {
    layout: "colorBlock",
    showAuthor: true,
    showPublishedDate: true,
    textAlignment: "left",
  },
  overlap: {
    layout: "overlap",
    enableGradientEffect: true,
    showAuthor: true,
    showPublishedDate: true,
    textAlignment: "left",
    imageWebUrl: "",
  },
};

export function registerTools(server, client, auth) {
  // ═══════════════════════════════════════════
  // SITE TOOLS
  // ═══════════════════════════════════════════

  server.tool(
    "search_sites",
    "Search for SharePoint sites by keyword",
    { query: z.string().describe("Search term for sites") },
    async ({ query }) => {
      const sites = await client.listSites(query);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              sites.map((s) => ({
                id: s.id,
                name: s.displayName,
                url: s.webUrl,
                description: s.description,
              })),
              null,
              2
            ),
          },
        ],
      };
    }
  );

  server.tool(
    "get_site_details",
    "Get details of a SharePoint site (ID, URL, description, lists)",
    { siteId: z.string().describe("Site ID (e.g. 'contoso.sharepoint.com,guid,guid')") },
    async ({ siteId }) => {
      const [site, lists] = await Promise.all([
        client.getSite(siteId),
        client.getSiteLists(siteId),
      ]);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              {
                site: {
                  id: site.id,
                  name: site.displayName,
                  url: site.webUrl,
                  description: site.description,
                  createdDateTime: site.createdDateTime,
                },
                lists: lists.map((l) => ({
                  id: l.id,
                  name: l.displayName,
                  template: l.list?.template,
                })),
              },
              null,
              2
            ),
          },
        ],
      };
    }
  );

  server.tool(
    "get_site_by_url",
    "Find a site by hostname and path",
    {
      hostname: z.string().describe("e.g. 'contoso.sharepoint.com'"),
      sitePath: z.string().describe("e.g. 'sites/marketing'"),
    },
    async ({ hostname, sitePath }) => {
      const site = await client.getSiteByUrl(hostname, sitePath);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              { id: site.id, name: site.displayName, url: site.webUrl },
              null,
              2
            ),
          },
        ],
      };
    }
  );

  server.tool(
    "list_my_sites",
    "List SharePoint sites the current user follows",
    {},
    async () => {
      try {
        const sites = await client.listSites("");
        return {
          content: [{
            type: "text",
            text: JSON.stringify(sites.map(s => ({
              id: s.id,
              name: s.displayName || s.name,
              url: s.webUrl,
              description: s.description,
            })), null, 2)
          }]
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: `Error listing followed sites: ${error.message}` }],
          isError: true,
        };
      }
    }
  );

  // ═══════════════════════════════════════════
  // PAGE TOOLS
  // ═══════════════════════════════════════════

  server.tool(
    "list_pages",
    "List all pages of a SharePoint site",
    { siteId: z.string() },
    async ({ siteId }) => {
      const pages = await client.listPages(siteId);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              pages.map((p) => ({
                id: p.id,
                name: p.name,
                title: p.title,
                webUrl: p.webUrl,
                pageLayout: p.pageLayout,
                lastModified: p.lastModifiedDateTime,
              })),
              null,
              2
            ),
          },
        ],
      };
    }
  );

  server.tool(
    "get_page",
    "Get the full content of a page including canvas layout and web parts",
    {
      siteId: z.string(),
      pageId: z.string().describe("Page ID"),
    },
    async ({ siteId, pageId }) => {
      const page = await client.getPage(siteId, pageId);
      return {
        content: [{ type: "text", text: JSON.stringify(page, null, 2) }],
      };
    }
  );

  server.tool(
    "create_page",
    "Create a new SharePoint page with optional layout and title area template",
    {
      siteId: z.string(),
      name: z.string().describe("File name (e.g. 'about-us.aspx')"),
      title: z.string().describe("Page title"),
      layout: z
        .enum(["article", "home", "repostPage"])
        .optional()
        .describe("Page layout"),
      titleLayout: z
        .enum(["plain", "imageAndTitle", "colorBlock", "overlap"])
        .optional()
        .describe("Title area design"),
      headerImageUrl: z.string().optional().describe("URL for header image"),
      autoPublish: z.boolean().optional().describe("Publish immediately?"),
    },
    async ({ siteId, name, title, layout, titleLayout, headerImageUrl, autoPublish }) => {
      const titleArea = titleLayout
        ? { ...TITLE_TEMPLATES[titleLayout] }
        : undefined;
      if (titleArea && headerImageUrl) {
        titleArea.imageWebUrl = headerImageUrl;
      }

      const page = await client.createPage(siteId, {
        name,
        title,
        layout,
        titleArea,
        headerImageUrl,
      });

      if (autoPublish) {
        await client.publishPage(siteId, page.id);
      }

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              {
                id: page.id,
                name: page.name,
                title: page.title,
                url: page.webUrl,
                status: autoPublish ? "published" : "draft",
              },
              null,
              2
            ),
          },
        ],
      };
    }
  );

  server.tool(
    "update_page",
    "Update page properties (title, title area, etc.)",
    {
      siteId: z.string(),
      pageId: z.string(),
      title: z.string().optional(),
      titleLayout: z
        .enum(["plain", "imageAndTitle", "colorBlock", "overlap"])
        .optional(),
      headerImageUrl: z.string().optional(),
      showComments: z.boolean().optional(),
      showRecommendedPages: z.boolean().optional(),
    },
    async ({ siteId, pageId, title, titleLayout, headerImageUrl, showComments, showRecommendedPages }) => {
      const updates = {};
      if (title) updates.title = title;
      if (showComments !== undefined) updates.showComments = showComments;
      if (showRecommendedPages !== undefined) updates.showRecommendedPages = showRecommendedPages;

      if (titleLayout) {
        updates.titleArea = { ...TITLE_TEMPLATES[titleLayout] };
        if (headerImageUrl) updates.titleArea.imageWebUrl = headerImageUrl;
      }

      await client.updatePage(siteId, pageId, updates);
      return {
        content: [{ type: "text", text: `Page ${pageId} updated.` }],
      };
    }
  );

  server.tool(
    "publish_page",
    "Publish a page",
    { siteId: z.string(), pageId: z.string() },
    async ({ siteId, pageId }) => {
      await client.publishPage(siteId, pageId);
      return {
        content: [{ type: "text", text: `Page ${pageId} published.` }],
      };
    }
  );

  server.tool(
    "delete_page",
    "Delete a page",
    { siteId: z.string(), pageId: z.string() },
    async ({ siteId, pageId }) => {
      await client.deletePage(siteId, pageId);
      return {
        content: [{ type: "text", text: `Page ${pageId} deleted.` }],
      };
    }
  );

  // ═══════════════════════════════════════════
  // CANVAS LAYOUT TOOLS (Sections)
  // ═══════════════════════════════════════════

  server.tool(
    "add_section",
    "Add a new section to a page",
    {
      siteId: z.string(),
      pageId: z.string(),
      sectionTemplate: z
        .enum([
          "fullWidth", "oneColumn", "twoColumns", "twoColumnsLeft",
          "twoColumnsRight", "threeColumns", "oneThirdLeft", "oneThirdRight",
        ])
        .describe("Layout template for the section"),
      emphasis: z
        .enum(["none", "neutral", "soft", "strong"])
        .optional()
        .describe("Background emphasis (color/contrast)"),
    },
    async ({ siteId, pageId, sectionTemplate, emphasis }) => {
      const template = { ...SECTION_TEMPLATES[sectionTemplate] };
      if (emphasis) template.emphasis = emphasis;

      const section = await client.addHorizontalSection(siteId, pageId, template);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              { message: "Section added", section },
              null,
              2
            ),
          },
        ],
      };
    }
  );

  server.tool(
    "get_page_layout",
    "Get the canvas layout of a page (sections, columns, web parts)",
    { siteId: z.string(), pageId: z.string() },
    async ({ siteId, pageId }) => {
      const [sections, webParts] = await Promise.all([
        client.getHorizontalSections(siteId, pageId),
        client.getPageWebParts(siteId, pageId),
      ]);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({ sections, webParts }, null, 2),
          },
        ],
      };
    }
  );

  // ═══════════════════════════════════════════
  // WEB PART TOOLS
  // ═══════════════════════════════════════════

  server.tool(
    "add_text_webpart",
    "Add a text web part (HTML) to a section",
    {
      siteId: z.string(),
      pageId: z.string(),
      sectionIndex: z.number().describe("Section index (1-based)"),
      columnIndex: z.number().describe("Column index (1-based)"),
      html: z.string().describe("HTML content (supports <h2>, <p>, <ul>, <a>, <strong>, <em>, etc.)"),
    },
    async ({ siteId, pageId, sectionIndex, columnIndex, html }) => {
      const webPart = WEB_PART_TEMPLATES.text(html);
      const result = await client.createWebPartInSection(
        siteId, pageId, { sectionIndex, columnIndex }, webPart
      );
      return {
        content: [
          { type: "text", text: JSON.stringify({ message: "Text web part added", id: result?.id }, null, 2) },
        ],
      };
    }
  );

  server.tool(
    "add_image_webpart",
    "Add an image web part to a section",
    {
      siteId: z.string(),
      pageId: z.string(),
      sectionIndex: z.number(),
      columnIndex: z.number(),
      imageUrl: z.string().describe("Image URL"),
      altText: z.string().optional(),
      caption: z.string().optional(),
    },
    async ({ siteId, pageId, sectionIndex, columnIndex, imageUrl, altText, caption }) => {
      const webPart = WEB_PART_TEMPLATES.image(imageUrl, altText, caption);
      const result = await client.createWebPartInSection(
        siteId, pageId, { sectionIndex, columnIndex }, webPart
      );
      return {
        content: [
          { type: "text", text: JSON.stringify({ message: "Image web part added", id: result?.id }, null, 2) },
        ],
      };
    }
  );

  server.tool(
    "add_spacer",
    "Add a spacer web part",
    {
      siteId: z.string(),
      pageId: z.string(),
      sectionIndex: z.number(),
      columnIndex: z.number(),
      height: z.number().optional().describe("Height in pixels (default: 60)"),
    },
    async ({ siteId, pageId, sectionIndex, columnIndex, height }) => {
      const webPart = WEB_PART_TEMPLATES.spacer(height);
      await client.createWebPartInSection(
        siteId, pageId, { sectionIndex, columnIndex }, webPart
      );
      return {
        content: [{ type: "text", text: "Spacer added" }],
      };
    }
  );

  server.tool(
    "add_divider",
    "Add a horizontal divider",
    {
      siteId: z.string(),
      pageId: z.string(),
      sectionIndex: z.number(),
      columnIndex: z.number(),
    },
    async ({ siteId, pageId, sectionIndex, columnIndex }) => {
      const webPart = WEB_PART_TEMPLATES.divider();
      await client.createWebPartInSection(
        siteId, pageId, { sectionIndex, columnIndex }, webPart
      );
      return {
        content: [{ type: "text", text: "Divider added" }],
      };
    }
  );

  server.tool(
    "add_custom_webpart",
    "Add a custom web part from a JSON definition (advanced)",
    {
      siteId: z.string(),
      pageId: z.string(),
      sectionIndex: z.number(),
      columnIndex: z.number(),
      webPartJson: z.string().describe("Web part JSON as string"),
    },
    async ({ siteId, pageId, sectionIndex, columnIndex, webPartJson }) => {
      const webPart = JSON.parse(webPartJson);
      const result = await client.createWebPartInSection(
        siteId, pageId, { sectionIndex, columnIndex }, webPart
      );
      return {
        content: [
          { type: "text", text: JSON.stringify({ message: "Web part added", id: result?.id }, null, 2) },
        ],
      };
    }
  );

  // ═══════════════════════════════════════════
  // NAVIGATION TOOLS
  // ═══════════════════════════════════════════

  server.tool(
    "get_navigation",
    "Get site navigation (Quick Launch or Top Navigation)",
    {
      siteUrl: z.string().describe("Full site URL, e.g. 'https://contoso.sharepoint.com/sites/marketing'"),
      navType: z.enum(["quick", "top"]).describe("Navigation type"),
    },
    async ({ siteUrl, navType }) => {
      const nav =
        navType === "top"
          ? await client.getTopNavigation(siteUrl)
          : await client.getNavigation(siteUrl);
      return {
        content: [{ type: "text", text: JSON.stringify(nav, null, 2) }],
      };
    }
  );

  server.tool(
    "add_navigation_link",
    "Add a navigation link (Quick Launch or Top Navigation)",
    {
      siteUrl: z.string(),
      navType: z.enum(["quick", "top"]),
      title: z.string().describe("Display name"),
      url: z.string().describe("Link URL"),
      isExternal: z.boolean().optional().describe("External link?"),
    },
    async ({ siteUrl, navType, title, url, isExternal }) => {
      await client.addNavigationNode(siteUrl, navType, { title, url, isExternal });
      return {
        content: [{ type: "text", text: `Navigation link "${title}" added` }],
      };
    }
  );

  server.tool(
    "remove_navigation_link",
    "Remove a navigation link",
    {
      siteUrl: z.string(),
      navType: z.enum(["quick", "top"]),
      nodeId: z.number().describe("Navigation node ID"),
    },
    async ({ siteUrl, navType, nodeId }) => {
      await client.deleteNavigationNode(siteUrl, navType, nodeId);
      return {
        content: [{ type: "text", text: `Navigation link ${nodeId} removed` }],
      };
    }
  );

  // ═══════════════════════════════════════════
  // BRANDING / ASSETS
  // ═══════════════════════════════════════════

  server.tool(
    "set_site_logo",
    "Set a SharePoint site logo",
    {
      siteUrl: z.string().describe("Full site URL"),
      logoUrl: z.string().describe("URL of the logo image"),
    },
    async ({ siteUrl, logoUrl }) => {
      await client.setSiteLogo(siteUrl, logoUrl);
      return {
        content: [{ type: "text", text: "Site logo updated" }],
      };
    }
  );

  server.tool(
    "upload_asset",
    "Upload a file (image, CSS, etc.) to Site Assets",
    {
      siteId: z.string(),
      fileName: z.string().describe("File name, e.g. 'hero-banner.jpg'"),
      folder: z.string().optional().describe("Folder path (default: 'SiteAssets')"),
      base64Content: z.string().describe("Base64-encoded file content"),
    },
    async ({ siteId, fileName, folder, base64Content }) => {
      const buffer = Buffer.from(base64Content, "base64");
      const targetFolder = folder || "SiteAssets";
      const result = await client.uploadFile(siteId, targetFolder, fileName, buffer);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              { message: "File uploaded", webUrl: result.webUrl, id: result.id },
              null,
              2
            ),
          },
        ],
      };
    }
  );

  // ═══════════════════════════════════════════
  // DESIGN HELPERS
  // ═══════════════════════════════════════════

  server.tool(
    "get_design_templates",
    "Get available design templates for sections, title areas, and web parts",
    {},
    async () => {
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              {
                sectionLayouts: Object.keys(SECTION_TEMPLATES),
                titleAreaLayouts: Object.keys(TITLE_TEMPLATES),
                webPartTypes: [
                  "text (HTML content)",
                  "image (image with alt text and caption)",
                  "hero (hero banner with items)",
                  "quickLinks (link collection)",
                  "spacer (whitespace)",
                  "divider (horizontal rule)",
                  "custom (any web part via JSON)",
                ],
                sectionEmphasis: ["none", "neutral", "soft", "strong"],
                tips: [
                  "Use 'overlap' or 'imageAndTitle' for visually appealing title area headers",
                  "Use 'strong' emphasis for colored sections as visual accents",
                  "Combine different column layouts for varied design",
                  "Use spacers between sections for whitespace",
                ],
              },
              null,
              2
            ),
          },
        ],
      };
    }
  );

  // ═══════════════════════════════════════════
  // AUTH / SESSION
  // ═══════════════════════════════════════════

  server.tool(
    "disconnect",
    "Disconnect from SharePoint and clear cached authentication. Use this to switch accounts or re-authenticate.",
    {},
    async () => {
      await auth.logout();
      return {
        content: [
          {
            type: "text",
            text: "Disconnected. Authentication cache cleared. You will be prompted to re-authenticate on the next request.",
          },
        ],
      };
    }
  );
}
