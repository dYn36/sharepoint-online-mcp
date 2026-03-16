# T01: Translate tool descriptions to English and add `list_my_sites`

## Description

Translate all 22 German MCP tool descriptions in `src/tools.js` to clear English. Add a new `list_my_sites` tool that returns the user's followed SharePoint sites. This is mechanical cleanup + one trivial tool addition.

## Steps

1. Open `src/tools.js` and locate all `server.tool(` registrations (lines ~118-694).
2. Translate each German description string to clear, concise English. Preserve the meaning. Examples:
   - `"Suche nach SharePoint-Sites anhand eines Suchbegriffs"` → `"Search for SharePoint sites by keyword"`
   - `"Zeigt Details einer SharePoint-Site (ID, URL, Beschreibung, Listen)"` → `"Get details of a SharePoint site (ID, URL, description, lists)"`
   - `"Listet alle Seiten einer SharePoint-Site"` → `"List all pages of a SharePoint site"`
   - `"Erstellt eine neue SharePoint-Seite..."` → `"Create a new SharePoint page..."`
   - All others follow the same pattern — short, imperative English
3. Do NOT change the `disconnect` tool description — it's already in English.
4. Add `list_my_sites` tool after the existing site tools group (`search_sites`, `get_site_details`, `get_site_by_url`) and before `list_pages`. Registration:
   ```javascript
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
   ```
5. Verify no German characters remain: `grep -P '[äöüÄÖÜß]' src/tools.js` should return empty.
6. Verify server starts: `node src/index.js` — should print startup message on stderr, no error.
7. Verify tool count: `grep -c 'server.tool(' src/tools.js` should return 24.

## Must-Haves

- All 22 German tool descriptions translated to English
- `list_my_sites` tool registered with no required parameters
- `list_my_sites` returns `{id, name, url, description}` array
- No regressions — existing tests pass, server starts clean

## Verification

```bash
grep -P '[äöüÄÖÜß]' src/tools.js  # must be empty
grep -c 'server.tool(' src/tools.js  # must return 24
node src/index.js 2>&1 &
PID=$!; sleep 1; kill $PID 2>/dev/null  # starts without error
node --test tests/auth.test.js && node --test tests/client.test.js  # no regressions
```

## Inputs

- `src/tools.js` — current state with 23 tool registrations, 22 with German descriptions, 1 English (`disconnect`)

## Expected Output

- `src/tools.js` modified: all descriptions in English, `list_my_sites` tool added (24 total tools)
- Zero German characters in the file
- Server starts cleanly, all existing tests pass
