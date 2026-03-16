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
    "Suche nach SharePoint-Sites anhand eines Suchbegriffs",
    { query: z.string().describe("Suchbegriff für Sites") },
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
    "Zeigt Details einer SharePoint-Site (ID, URL, Beschreibung, Listen)",
    { siteId: z.string().describe("Site-ID (z.B. 'contoso.sharepoint.com,guid,guid')") },
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
    "Findet eine Site anhand von Hostname und Pfad",
    {
      hostname: z.string().describe("z.B. 'contoso.sharepoint.com'"),
      sitePath: z.string().describe("z.B. 'sites/marketing'"),
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

  // ═══════════════════════════════════════════
  // PAGE TOOLS
  // ═══════════════════════════════════════════

  server.tool(
    "list_pages",
    "Listet alle Seiten einer SharePoint-Site",
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
    "Zeigt den vollständigen Inhalt einer Seite inkl. Canvas-Layout und Web Parts",
    {
      siteId: z.string(),
      pageId: z.string().describe("Page-ID"),
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
    "Erstellt eine neue SharePoint-Seite mit optionalem Layout und Title-Area-Template",
    {
      siteId: z.string(),
      name: z.string().describe("Dateiname (z.B. 'about-us.aspx')"),
      title: z.string().describe("Seitentitel"),
      layout: z
        .enum(["article", "home", "repostPage"])
        .optional()
        .describe("Seitenlayout"),
      titleLayout: z
        .enum(["plain", "imageAndTitle", "colorBlock", "overlap"])
        .optional()
        .describe("Title-Area-Design"),
      headerImageUrl: z.string().optional().describe("URL für Header-Bild"),
      autoPublish: z.boolean().optional().describe("Sofort veröffentlichen?"),
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
    "Aktualisiert Seiteneigenschaften (Titel, Title Area, etc.)",
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
        content: [{ type: "text", text: `✅ Seite ${pageId} aktualisiert.` }],
      };
    }
  );

  server.tool(
    "publish_page",
    "Veröffentlicht eine Seite",
    { siteId: z.string(), pageId: z.string() },
    async ({ siteId, pageId }) => {
      await client.publishPage(siteId, pageId);
      return {
        content: [{ type: "text", text: `✅ Seite ${pageId} veröffentlicht.` }],
      };
    }
  );

  server.tool(
    "delete_page",
    "Löscht eine Seite",
    { siteId: z.string(), pageId: z.string() },
    async ({ siteId, pageId }) => {
      await client.deletePage(siteId, pageId);
      return {
        content: [{ type: "text", text: `🗑️ Seite ${pageId} gelöscht.` }],
      };
    }
  );

  // ═══════════════════════════════════════════
  // CANVAS LAYOUT TOOLS (Sections)
  // ═══════════════════════════════════════════

  server.tool(
    "add_section",
    "Fügt einen neuen Abschnitt (Section) zu einer Seite hinzu",
    {
      siteId: z.string(),
      pageId: z.string(),
      sectionTemplate: z
        .enum([
          "fullWidth", "oneColumn", "twoColumns", "twoColumnsLeft",
          "twoColumnsRight", "threeColumns", "oneThirdLeft", "oneThirdRight",
        ])
        .describe("Layout-Template für die Section"),
      emphasis: z
        .enum(["none", "neutral", "soft", "strong"])
        .optional()
        .describe("Hintergrund-Emphasis (Farbe/Kontrast)"),
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
              { message: "✅ Section hinzugefügt", section },
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
    "Zeigt das Canvas-Layout einer Seite (Sections, Columns, Web Parts)",
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
    "Fügt einen Text-Web-Part (HTML) in eine Sektion ein",
    {
      siteId: z.string(),
      pageId: z.string(),
      sectionIndex: z.number().describe("Section-Index (1-basiert)"),
      columnIndex: z.number().describe("Column-Index (1-basiert)"),
      html: z.string().describe("HTML-Inhalt (unterstützt <h2>, <p>, <ul>, <a>, <strong>, <em>, etc.)"),
    },
    async ({ siteId, pageId, sectionIndex, columnIndex, html }) => {
      const webPart = WEB_PART_TEMPLATES.text(html);
      const result = await client.createWebPartInSection(
        siteId, pageId, { sectionIndex, columnIndex }, webPart
      );
      return {
        content: [
          { type: "text", text: JSON.stringify({ message: "✅ Text hinzugefügt", id: result?.id }, null, 2) },
        ],
      };
    }
  );

  server.tool(
    "add_image_webpart",
    "Fügt ein Bild in eine Sektion ein",
    {
      siteId: z.string(),
      pageId: z.string(),
      sectionIndex: z.number(),
      columnIndex: z.number(),
      imageUrl: z.string().describe("URL des Bildes"),
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
          { type: "text", text: JSON.stringify({ message: "✅ Bild hinzugefügt", id: result?.id }, null, 2) },
        ],
      };
    }
  );

  server.tool(
    "add_spacer",
    "Fügt einen Abstandshalter (Spacer) ein",
    {
      siteId: z.string(),
      pageId: z.string(),
      sectionIndex: z.number(),
      columnIndex: z.number(),
      height: z.number().optional().describe("Höhe in Pixel (Standard: 60)"),
    },
    async ({ siteId, pageId, sectionIndex, columnIndex, height }) => {
      const webPart = WEB_PART_TEMPLATES.spacer(height);
      await client.createWebPartInSection(
        siteId, pageId, { sectionIndex, columnIndex }, webPart
      );
      return {
        content: [{ type: "text", text: "✅ Spacer hinzugefügt" }],
      };
    }
  );

  server.tool(
    "add_divider",
    "Fügt eine horizontale Trennlinie ein",
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
        content: [{ type: "text", text: "✅ Trennlinie hinzugefügt" }],
      };
    }
  );

  server.tool(
    "add_custom_webpart",
    "Fügt einen beliebigen Web Part per JSON-Definition ein (für fortgeschrittene Nutzer)",
    {
      siteId: z.string(),
      pageId: z.string(),
      sectionIndex: z.number(),
      columnIndex: z.number(),
      webPartJson: z.string().describe("Web Part JSON als String"),
    },
    async ({ siteId, pageId, sectionIndex, columnIndex, webPartJson }) => {
      const webPart = JSON.parse(webPartJson);
      const result = await client.createWebPartInSection(
        siteId, pageId, { sectionIndex, columnIndex }, webPart
      );
      return {
        content: [
          { type: "text", text: JSON.stringify({ message: "✅ Web Part hinzugefügt", id: result?.id }, null, 2) },
        ],
      };
    }
  );

  // ═══════════════════════════════════════════
  // NAVIGATION TOOLS
  // ═══════════════════════════════════════════

  server.tool(
    "get_navigation",
    "Zeigt die Navigation einer Site (Quick Launch oder Top Nav)",
    {
      siteUrl: z.string().describe("Volle Site-URL, z.B. 'https://contoso.sharepoint.com/sites/marketing'"),
      navType: z.enum(["quick", "top"]).describe("Navigationstyp"),
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
    "Fügt einen Navigationslink hinzu (Quick Launch oder Top Nav)",
    {
      siteUrl: z.string(),
      navType: z.enum(["quick", "top"]),
      title: z.string().describe("Anzeigename"),
      url: z.string().describe("Link-URL"),
      isExternal: z.boolean().optional().describe("Externer Link?"),
    },
    async ({ siteUrl, navType, title, url, isExternal }) => {
      await client.addNavigationNode(siteUrl, navType, { title, url, isExternal });
      return {
        content: [{ type: "text", text: `✅ Navigation "${title}" hinzugefügt` }],
      };
    }
  );

  server.tool(
    "remove_navigation_link",
    "Entfernt einen Navigationslink",
    {
      siteUrl: z.string(),
      navType: z.enum(["quick", "top"]),
      nodeId: z.number().describe("ID des Nav-Knotens"),
    },
    async ({ siteUrl, navType, nodeId }) => {
      await client.deleteNavigationNode(siteUrl, navType, nodeId);
      return {
        content: [{ type: "text", text: `🗑️ Navigation ${nodeId} entfernt` }],
      };
    }
  );

  // ═══════════════════════════════════════════
  // BRANDING / ASSETS
  // ═══════════════════════════════════════════

  server.tool(
    "set_site_logo",
    "Setzt das Logo einer SharePoint-Site",
    {
      siteUrl: z.string().describe("Volle Site-URL"),
      logoUrl: z.string().describe("URL zum Logo-Bild"),
    },
    async ({ siteUrl, logoUrl }) => {
      await client.setSiteLogo(siteUrl, logoUrl);
      return {
        content: [{ type: "text", text: "✅ Site-Logo aktualisiert" }],
      };
    }
  );

  server.tool(
    "upload_asset",
    "Lädt eine Datei (Bild, CSS, etc.) in die Site Assets hoch",
    {
      siteId: z.string(),
      fileName: z.string().describe("Dateiname z.B. 'hero-banner.jpg'"),
      folder: z.string().optional().describe("Ordnerpfad (Standard: 'SiteAssets')"),
      base64Content: z.string().describe("Base64-kodierter Dateiinhalt"),
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
              { message: "✅ Datei hochgeladen", webUrl: result.webUrl, id: result.id },
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
    "Zeigt verfügbare Design-Templates für Sections, Title Areas und Web Parts",
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
                  "text (HTML-Inhalt)",
                  "image (Bild mit Alt-Text und Caption)",
                  "hero (Hero-Banner mit Items)",
                  "quickLinks (Link-Sammlung)",
                  "spacer (Abstandshalter)",
                  "divider (Trennlinie)",
                  "custom (beliebiger Web Part per JSON)",
                ],
                sectionEmphasis: ["none", "neutral", "soft", "strong"],
                tips: [
                  "Verwende 'overlap' oder 'imageAndTitle' für Title Areas mit visuell ansprechendem Header",
                  "Nutze 'strong' Emphasis für farbige Sections als visuelle Akzente",
                  "Kombiniere verschiedene Column-Layouts für abwechslungsreiches Design",
                  "Nutze Spacer zwischen Sections für Whitespace",
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
