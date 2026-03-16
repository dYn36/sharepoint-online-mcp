# SharePoint MCP Server

Ein MCP-Server (Model Context Protocol) für SharePoint Online, der es ermöglicht, Sites direkt über Claude zu bearbeiten und zu designen – **ohne Admin-Rechte**.

## ✨ Features

| Bereich | Tools |
|---------|-------|
| **Sites** | Suchen, Details anzeigen, per URL finden |
| **Seiten** | Erstellen, bearbeiten, veröffentlichen, löschen |
| **Layout** | Sections hinzufügen (1-3 Spalten, Full Width, etc.) |
| **Web Parts** | Text (HTML), Bilder, Spacer, Trennlinien, Custom |
| **Navigation** | Quick Launch & Top Nav lesen, Links hinzufügen/entfernen |
| **Branding** | Site-Logo setzen, Assets hochladen |
| **Design-Hilfe** | Template-Übersicht für Layouts & Web Parts |

## 🔧 Einrichtung (ohne Admin-Rechte!)

### Schritt 1: Azure AD App registrieren

> Du brauchst **keine** Admin-Rechte dafür – jeder User kann in den meisten Organisationen eigene App-Registrierungen erstellen.

1. Öffne [portal.azure.com](https://portal.azure.com)
2. Gehe zu **Azure Active Directory** → **App registrations**
3. Klicke **"New registration"**
4. Einstellungen:
   - **Name:** `SharePoint MCP` (oder beliebig)
   - **Supported account types:** `Accounts in this organizational directory only`
   - **Redirect URI:** leer lassen
5. Klicke **"Register"**
6. Kopiere die **Application (client) ID** und die **Directory (tenant) ID**

### Schritt 2: Authentication konfigurieren

1. Im App-Menü: **Authentication**
2. Scrolle runter zu **"Advanced settings"**
3. Setze **"Allow public client flows"** auf **Yes**
4. **Speichern**

### Schritt 3: API-Berechtigungen

1. Im App-Menü: **API permissions** → **Add a permission**
2. Wähle **Microsoft Graph** → **Delegated permissions**
3. Füge hinzu:
   - `Sites.ReadWrite.All` – Sites lesen und schreiben
   - `Sites.Manage.All` – Sites verwalten
   - `User.Read` – User-Profil lesen
4. **Kein Admin Consent nötig!** Diese delegierten Berechtigungen funktionieren mit deinem User-Account.

### Schritt 4: Installation

```bash
# Repository klonen / kopieren
cd sharepoint-mcp
npm install
```

### Schritt 5: Umgebungsvariablen

```bash
export SHAREPOINT_CLIENT_ID="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
export SHAREPOINT_TENANT_ID="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
```

## 🔌 In Claude Desktop einbinden

Füge dies in deine `claude_desktop_config.json` ein:

**macOS:** `~/Library/Application Support/Claude/claude_desktop_config.json`
**Windows:** `%APPDATA%\Claude\claude_desktop_config.json`

```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "node",
      "args": ["/PFAD/ZU/sharepoint-mcp/src/index.js"],
      "env": {
        "SHAREPOINT_CLIENT_ID": "deine-client-id",
        "SHAREPOINT_TENANT_ID": "deine-tenant-id"
      }
    }
  }
}
```

## 🔐 Login-Flow

Beim ersten Aufruf eines Tools wirst du im Terminal aufgefordert:

```
🔐 SharePoint Login erforderlich!
   Öffne: https://microsoft.com/devicelogin
   Code:  ABCD-EFGH
```

1. Öffne die URL im Browser
2. Gib den Code ein
3. Melde dich mit deinem Microsoft-Account an
4. Das Token wird lokal gecacht (`~/.sharepoint-mcp-cache.json`)

Danach funktioniert alles automatisch, bis das Token abläuft.

## 📖 Nutzungsbeispiele

### Site finden
```
"Suche meine Marketing-SharePoint-Site"
→ search_sites("marketing")
```

### Neue Seite erstellen
```
"Erstelle eine 'Über uns'-Seite mit einem schönen Header-Bild"
→ create_page mit titleLayout: "imageAndTitle"
```

### Seite designen
```
"Füge eine zweispaltige Section hinzu mit Text links und Bild rechts"
→ add_section(twoColumns) → add_text_webpart(col 1) → add_image_webpart(col 2)
```

### Navigation bearbeiten
```
"Füge einen Link zur neuen Seite in die Quick Launch Navigation"
→ add_navigation_link(navType: "quick", ...)
```

## 🏗️ Verfügbare Tools

### Sites
- `search_sites` – Sites suchen
- `get_site_details` – Site-Details mit Listen
- `get_site_by_url` – Site per Hostname/Pfad finden

### Seiten
- `list_pages` – Alle Seiten auflisten
- `get_page` – Seiteninhalt lesen
- `create_page` – Neue Seite erstellen
- `update_page` – Seiteneigenschaften ändern
- `publish_page` – Seite veröffentlichen
- `delete_page` – Seite löschen

### Layout & Web Parts
- `get_page_layout` – Canvas-Layout anzeigen
- `add_section` – Section hinzufügen (8 Layouts)
- `add_text_webpart` – HTML-Text einfügen
- `add_image_webpart` – Bild einfügen
- `add_spacer` – Abstandshalter
- `add_divider` – Trennlinie
- `add_custom_webpart` – Beliebiger Web Part (JSON)

### Navigation
- `get_navigation` – Navigation lesen
- `add_navigation_link` – Link hinzufügen
- `remove_navigation_link` – Link entfernen

### Branding
- `set_site_logo` – Logo setzen
- `upload_asset` – Dateien hochladen

### Hilfe
- `get_design_templates` – Template-Übersicht

## ⚠️ Einschränkungen

- **Keine Admin-Features**: Site Designs, Hub Sites, Tenant-Themes sind Admin-only
- **Nur deine Berechtigungen**: Du kannst nur auf Sites zugreifen, auf die du auch im Browser Zugriff hast
- **Graph Beta API**: Einige Page-Endpoints nutzen die Beta API, die sich ändern kann
- **Web Part-Typen**: Nicht alle Web Part-Typen sind über die API erstellbar (z.B. Power BI)
- **SP REST API**: Navigation-Endpoints nutzen die klassische SharePoint REST API

## 🛠️ Troubleshooting

| Problem | Lösung |
|---------|--------|
| `AADSTS65001` | Dein Tenant erlaubt keine User-App-Registrierung → IT fragen |
| `403 Forbidden` | Keine Berechtigung auf die Site → Zugriff im Browser prüfen |
| `Token expired` | Cache löschen: `rm ~/.sharepoint-mcp-cache.json` |
| Pages API fehlt | Beta API nur für moderne Sites, nicht für klassische Sites |
