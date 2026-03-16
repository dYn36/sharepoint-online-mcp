# SharePoint Online MCP Server

MCP server for SharePoint Online — zero-config, no Azure Portal, no env vars.

## Quick Start

Install and run:

```bash
npx sharepoint-online-mcp
```

Add to Claude Desktop (`claude_desktop_config.json`):

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

No environment variables needed. That's the entire setup.

## How It Works

- Uses the Microsoft Office well-known client ID — no Azure app registration needed
- On first use, prompts for device code authentication (open a URL, enter a code)
- Auto-discovers your tenant ID from the SharePoint URL
- Caches tokens in `~/.sharepoint-mcp-cache.json` — persists across restarts

## Usage

Start by asking Claude to connect to your SharePoint site:

> "Connect to https://contoso.sharepoint.com/sites/marketing"

Claude will handle the authentication prompts. Once connected, use any tool — search sites, create pages, manage navigation, and more.

## Tools

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

## Requirements

- Node.js >= 18
- A Microsoft 365 / SharePoint Online account

## Known Limitations

- Some organizations block device code authentication via Conditional Access policies. If you see error `AADSTS50076` or `AADSTS53003`, contact your IT administrator.
- Only works with SharePoint Online (Microsoft 365), not on-premises SharePoint Server.
- Permissions depend on your Microsoft 365 account — you can only access sites you have permission to.

## Troubleshooting

| Problem | Solution |
|---------|----------|
| Authentication failed | Run the `disconnect` tool and try again |
| Tenant discovery failed | Check that the SharePoint URL is correct and accessible |
| Access denied (403) | Your account may not have permission for this site or operation |
| Resource not found (404) | Verify the site URL, page ID, or resource exists |
| Token issues | Delete `~/.sharepoint-mcp-cache.json` and re-authenticate |

## License

MIT
