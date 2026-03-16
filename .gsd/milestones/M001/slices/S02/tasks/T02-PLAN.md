# T02: Add `connect_to_site` tool with URL parsing and tests

## Description

Add the `connect_to_site` MCP tool — the primary new capability in S02. Takes a single SharePoint URL string, parses hostname and site path, discovers the tenant ID, and resolves the site via Graph API. Also create the slice's test file `tests/tools.test.js` covering URL parsing edge cases and tool registration.

## Steps

1. **Extract URL parsing as a testable function.** Add and export a `parseSharePointUrl(url)` function in `src/tools.js` (at the top, before `registerTools`). This function:
   - Takes a single string `url`
   - Parses with `new URL(url)` — catch TypeError for invalid URLs, throw with `"Invalid SharePoint URL: {url}"`
   - Extracts `hostname` from the parsed URL
   - Extracts site path from `pathname`:
     - Split pathname by `/`, filter empty segments
     - If first segment is `sites` or `teams`, site path = `{segment[0]}/{segment[1]}` (e.g. `sites/marketing`)
     - If no segments or first segment is not `sites`/`teams`, site path = `` (empty string — root site)
     - Ignore any segments beyond the first two (handles `/sites/marketing/SitePages/Home.aspx`)
   - Returns `{ hostname, sitePath }` — e.g. `{ hostname: "contoso.sharepoint.com", sitePath: "sites/marketing" }`

2. **Add `connect_to_site` tool registration.** Place it in the site tools group (after `get_site_by_url`, before `list_my_sites`). Import `discoverTenantId` from `../auth.js` at the top of the `registerTools` function or at the module level. Registration:
   ```javascript
   server.tool(
     "connect_to_site",
     "Connect to a SharePoint site from its URL. Automatically discovers the tenant and resolves the site. Use this as the starting point when a user provides a SharePoint URL.",
     { url: z.string().url().describe("Full SharePoint site URL, e.g. https://contoso.sharepoint.com/sites/marketing") },
     async ({ url }) => {
       try {
         const { hostname, sitePath } = parseSharePointUrl(url);
         const tenantId = await discoverTenantId(hostname);
         const site = sitePath
           ? await client.getSiteByUrl(hostname, sitePath)
           : await client.getSiteByUrl(hostname, "/");
         return {
           content: [{
             type: "text",
             text: JSON.stringify({
               id: site.id,
               name: site.displayName || site.name,
               url: site.webUrl,
               description: site.description,
               tenantId,
             }, null, 2)
           }]
         };
       } catch (error) {
         return {
           content: [{ type: "text", text: `Error connecting to site: ${error.message}` }],
           isError: true,
         };
       }
     }
   );
   ```
   
   **Important:** `discoverTenantId` import — add at the module top level:
   ```javascript
   import { discoverTenantId } from "./auth.js";
   ```
   This is the first cross-module import in tools.js beyond `zod`. The file currently imports only `{ z } from "zod"`.

3. **Handle root site edge case.** When `sitePath` is empty (root site URL like `https://contoso.sharepoint.com`), pass `"/"` to `getSiteByUrl`. The Graph API endpoint `sites/{hostname}:/` resolves to the root site. Check what `getSiteByUrl` does with an empty vs `/` path — look at `src/client.js` line ~88.

4. **Create `tests/tools.test.js`.** Test structure:
   ```javascript
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
   ```

5. Run all tests:
   ```bash
   node --test tests/tools.test.js
   node --test tests/auth.test.js
   node --test tests/client.test.js
   ```

6. Verify final tool count: `grep -c 'server.tool(' src/tools.js` → 25

7. Verify server starts: `node src/index.js` starts without error on stderr.

## Must-Haves

- `parseSharePointUrl` exported and handles: standard sites, root site, trailing slashes, subpage paths, teams prefix, invalid URLs
- `connect_to_site` tool registered with `url` parameter (z.string().url())
- `connect_to_site` calls `discoverTenantId` then `getSiteByUrl`, returns site details + tenant ID
- `connect_to_site` returns actionable error for invalid URL input
- `discoverTenantId` imported from `./auth.js` — first cross-module import in tools.js
- `tests/tools.test.js` created with 7+ test cases for `parseSharePointUrl`
- All existing tests (18) still pass — no regressions

## Verification

```bash
node --test tests/tools.test.js  # all new tests pass
node --test tests/auth.test.js && node --test tests/client.test.js  # 18 existing tests pass
grep -c 'server.tool(' src/tools.js  # returns 25
node src/index.js 2>&1 &
PID=$!; sleep 1; kill $PID 2>/dev/null  # starts without error
```

## Inputs

- `src/tools.js` — state after T01 (24 tools, all English descriptions, `list_my_sites` added)
- `src/auth.js` — exports `discoverTenantId(domain)` (from S01)
- `src/client.js` — has `getSiteByUrl(hostname, sitePath)` and `listSites(search)` (from S01)
- `tests/auth.test.js` — 12 tests (from S01)
- `tests/client.test.js` — 6 tests (from S01)

## Expected Output

- `src/tools.js` modified: `parseSharePointUrl` exported, `connect_to_site` tool registered, `discoverTenantId` imported from `./auth.js` (25 total tools)
- `tests/tools.test.js` created: 7+ tests for URL parsing, all passing
- All 18 existing tests still pass
- Server starts cleanly

## Observability Impact

- `connect_to_site` error paths include the original URL in the error message for diagnostic context
- Tenant discovery errors propagate from `discoverTenantId` which already includes domain and HTTP status
- No new stderr logging beyond what `discoverTenantId` already emits
