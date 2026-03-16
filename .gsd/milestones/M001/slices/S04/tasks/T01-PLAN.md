# T01: Make package npm-publish-ready

## Description

The package needs to be publishable to npm so `npx sharepoint-online-mcp` works. Currently missing: `files` whitelist, `engines`, `keywords`, `license`, `repository` in package.json. No LICENSE file exists. No `.npmignore`. `src/index.js` lacks executable permission (currently `-rw-r--r--`).

## Steps

1. **Edit `package.json`** — add these fields (keep all existing fields intact):
   - `"license": "MIT"`
   - `"engines": { "node": ">=18" }`
   - `"files": ["src/", "README.md", "LICENSE"]`
   - `"keywords": ["sharepoint", "mcp", "model-context-protocol", "sharepoint-online", "claude", "ai"]`
   - `"repository": { "type": "git", "url": "https://github.com/daniellaurin/sharepoint-online-mcp.git" }`

2. **Create `LICENSE`** — MIT license, year 2026, copyright holder "Daniel Laurin". Standard MIT text.

3. **Create `.npmignore`** — exclude everything that shouldn't be in the published package:
   ```
   tests/
   .gsd/
   .git/
   .gitignore
   .npmignore
   .env
   .env.*
   node_modules/
   *.test.js
   ```

4. **Fix executable permission** — run `chmod +x src/index.js`. The file already has `#!/usr/bin/env node` shebang on line 1.

5. **Verify with `npm pack --dry-run`** — output must include ONLY:
   - `package.json`
   - `README.md`
   - `LICENSE`
   - `src/index.js`
   - `src/auth.js`
   - `src/auth-cli.js`
   - `src/client.js`
   - `src/tools.js`
   
   Must NOT include: `tests/`, `.gsd/`, `.git/`, `.env`, `node_modules/`

6. **Verify `npx .` starts the server** — run `npx . 2>&1 | head -1` from the project root. Expected output: `🚀 SharePoint Online MCP Server started (stdio)` (the process will hang on stdin since it's an MCP server — just capture the first line and kill it).

7. **Verify regression** — run `node --test tests/*.test.js` — all 56 tests must still pass.

## Must-Haves

- `files` field in package.json whitelists only `src/`, `README.md`, `LICENSE`
- `engines` field specifies `node >= 18` (global fetch dependency)
- MIT LICENSE file exists
- `src/index.js` is executable
- `npm pack --dry-run` includes only intended files

## Inputs

- Current `package.json` has: `name`, `version`, `description`, `type`, `main`, `bin`, `scripts`, `dependencies` — all correct. Only metadata fields are missing.
- `src/index.js` already has `#!/usr/bin/env node` shebang at line 1.
- No LICENSE file exists yet.
- No `.npmignore` exists yet.

## Expected Output

- `package.json` with all npm metadata fields populated
- `LICENSE` file (MIT, 2026)
- `.npmignore` file
- `src/index.js` with executable permission
- `npm pack --dry-run` shows clean file list
- `npx .` starts server successfully
- All 56 existing tests still pass

## Verification

```bash
npm pack --dry-run 2>&1
npx . 2>&1 | head -1
node --test tests/*.test.js
```
