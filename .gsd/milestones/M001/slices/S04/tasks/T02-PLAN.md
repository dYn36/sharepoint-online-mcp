# T02: Rewrite README for zero-config workflow

## Description

The README is entirely in German and documents the old workflow (Azure AD app registration, env vars, manual tenant ID). This is now wrong — the server uses zero-config auth with a well-known client ID, device code flow, and auto tenant discovery. Full rewrite in English.

## Steps

1. **Rewrite `README.md`** with this structure:

   **Title section:**
   - `# SharePoint Online MCP Server`
   - One-liner: MCP server for SharePoint Online — zero-config, no Azure Portal, no env vars.

   **Quick Start:**
   - Install: `npx sharepoint-online-mcp`
   - Claude Desktop config JSON (in `claude_desktop_config.json`):
     ```json
     {
       "mcpServers": {
         "sharepoint": {
           "command": "npx",
           "args": ["-y", "sharepoint-online-mcp"]
         }
       }
     }
     ```
   - No env vars needed. That's the entire setup.

   **How It Works:**
   - Uses Microsoft Office well-known client ID — no app registration needed
   - On first use, prompts for device code authentication (open URL, enter code)
   - Auto-discovers tenant ID from SharePoint URL
   - Caches tokens in `~/.sharepoint-mcp-cache.json` — persists across restarts

   **Usage:**
   - Start by asking Claude to connect: "Connect to https://contoso.sharepoint.com/sites/marketing"
   - Claude will handle authentication prompts
   - Then use any tool: search sites, create pages, manage navigation, etc.

   **Tool Reference Table:**
   - Organized by category with these exact tool names (from `src/tools.js`):

   | Category | Tool | Description |
   |----------|------|-------------|
   | **Sites** | `search_sites` | Search for SharePoint sites by keyword |
   | | `get_site_details` | Get detailed information about a site |
   | | `get_site_by_url` | Resolve a site from its SharePoint URL |
   | | `connect_to_site` | Connect to a site (auto tenant discovery + auth) |
   | | `list_my_sites` | List sites you follow |
   | **Pages** | `list_pages` | List pages in a site |
   | | `get_page` | Get page content and metadata |
   | | `create_page` | Create a new modern page |
   | | `update_page` | Update page title or description |
   | | `publish_page` | Publish a draft page |
   | | `delete_page` | Delete a page |
   | **Layout** | `add_section` | Add a section with column layout |
   | | `get_page_layout` | Get the layout structure of a page |
   | **Web Parts** | `add_text_webpart` | Add a text/HTML web part |
   | | `add_image_webpart` | Add an image web part |
   | | `add_spacer` | Add a spacer web part |
   | | `add_divider` | Add a divider web part |
   | | `add_custom_webpart` | Add any web part by type ID |
   | **Navigation** | `get_navigation` | Get Quick Launch or Top Navigation |
   | | `add_navigation_link` | Add a navigation link |
   | | `remove_navigation_link` | Remove a navigation link |
   | **Branding** | `set_site_logo` | Set the site logo |
   | | `upload_asset` | Upload a file to Site Assets |
   | **Utility** | `get_design_templates` | List available site design templates |
   | | `disconnect` | Clear auth tokens and disconnect |

   **Requirements:**
   - Node.js >= 18
   - A Microsoft 365 / SharePoint Online account

   **Known Limitations:**
   - Some organizations block device code authentication via Conditional Access policies. If you see error `AADSTS50076` or `AADSTS53003`, contact your IT administrator.
   - Only works with SharePoint Online (Microsoft 365), not on-premises SharePoint Server.
   - Permissions depend on your Microsoft 365 account — you can only access sites you have permission to.

   **Troubleshooting:**
   - "Authentication failed" → Run `disconnect` tool and try again
   - "Tenant discovery failed" → Check that the SharePoint URL is correct and accessible
   - "Access denied (403)" → Your account may not have permission for this site or operation
   - "Resource not found (404)" → Verify the site URL, page ID, or resource exists
   - Token issues → Delete `~/.sharepoint-mcp-cache.json` and re-authenticate

   **License:**
   - MIT

2. **Verify** — no German fragments remain, no references to env vars / `SHAREPOINT_TENANT_ID` / `AZURE_CLIENT_ID` / app registration, all 25 tool names are present and match `src/tools.js`.

## Must-Haves

- README is entirely in English
- Zero references to Azure app registration, env vars, or manual tenant ID configuration
- Claude Desktop config example has NO env block — just `command` and `args`
- All 25 tool names match exactly what's registered in `src/tools.js`
- Known limitations section mentions Conditional Access
- Troubleshooting section covers common error scenarios

## Inputs

- Tool names (all 25, from `src/tools.js`): `search_sites`, `get_site_details`, `get_site_by_url`, `connect_to_site`, `list_my_sites`, `list_pages`, `get_page`, `create_page`, `update_page`, `publish_page`, `delete_page`, `add_section`, `get_page_layout`, `add_text_webpart`, `add_image_webpart`, `add_spacer`, `add_divider`, `add_custom_webpart`, `get_navigation`, `add_navigation_link`, `remove_navigation_link`, `set_site_logo`, `upload_asset`, `get_design_templates`, `disconnect`
- Package name: `sharepoint-online-mcp`
- Auth mechanism: Microsoft Office well-known client ID `d3590ed6-52b3-4102-aeff-aad2292ab01c`, device code flow
- Token cache: `~/.sharepoint-mcp-cache.json`
- Node requirement: >= 18 (global fetch)

## Expected Output

- `README.md` — complete English README for zero-config SharePoint MCP server
- No German text anywhere
- All 25 tool names present and correct

## Observability Impact

- **README accuracy signal:** `grep -c` for all 25 tool names against `src/tools.js` — detects drift if tools are added/removed later. Run after any tool registration change.
- **No stale references:** `grep -i "SHAREPOINT_TENANT_ID\|AZURE_CLIENT_ID\|\.env"` against README should always return 0 matches. Any match means old config leaked back in.
- **Language check:** `grep -i` with German keyword list against README catches accidental German re-introduction (note: "Branding" is a false positive — it's an English category label).

## Verification

```bash
# No German fragments
grep -i "einrichtung\|Umgebung\|Berechtigung\|Schritt\|bearbeiten\|erstellen\|verfügbar\|Seiten\|Branding\|Hilfe\|Einschränk\|ermöglicht\|konfigurieren" README.md | wc -l
# Should be 0

# All tool names present
for tool in search_sites get_site_details get_site_by_url connect_to_site list_my_sites list_pages get_page create_page update_page publish_page delete_page add_section get_page_layout add_text_webpart add_image_webpart add_spacer add_divider add_custom_webpart get_navigation add_navigation_link remove_navigation_link set_site_logo upload_asset get_design_templates disconnect; do
  grep -q "$tool" README.md || echo "MISSING: $tool"
done

# No env var references
grep -i "SHAREPOINT_TENANT_ID\|AZURE_CLIENT_ID\|SHAREPOINT_CLIENT_SECRET\|\.env" README.md | wc -l
# Should be 0
```
