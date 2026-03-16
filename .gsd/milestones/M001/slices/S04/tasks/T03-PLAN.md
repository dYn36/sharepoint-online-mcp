# T03: Wrap auth and client errors with actionable guidance

## Description

R015 — error messages currently pass through raw Microsoft error text. In a zero-config setup, users have no config knobs to adjust, so errors must be self-explanatory. This task wraps known AADSTS codes in auth.js and HTTP status codes in client.js with actionable guidance, then adds tests for the new error messages.

## Steps

1. **Enhance `src/auth.js` error handling** — in the `acquireTokenByDeviceCode` flow (or wherever the MSAL device code call is made), wrap the catch block to detect and remap common AADSTS error codes:

   - `AADSTS50076` or `AADSTS53003` → `"Authentication blocked: Your organization's Conditional Access policy does not allow device code flow. Contact your IT administrator or try from a different network."`
   - `AADSTS700016` → `"Authentication error: Application not recognized. This is unexpected — please report this issue."`
   - `AADSTS50059` → `"Tenant not found: Could not find a Microsoft 365 tenant for this domain. Check that the SharePoint URL is correct."`
   - Any other AADSTS error → `"Authentication failed (CODE): ORIGINAL_MESSAGE. Try using the 'disconnect' tool and authenticating again."`
   - Non-AADSTS errors → `"Authentication failed: ORIGINAL_MESSAGE. Try using the 'disconnect' tool and authenticating again."`

   Implementation approach: In the catch block, check if `error.message` or `error.errorCode` contains the AADSTS code patterns. MSAL errors have an `errorCode` property. Keep the original error message in the rethrown error for debugging but prepend the actionable guidance.

   **Important:** The existing `discoverTenantId` error messages (lines 45-67 in auth.js) are already reasonably actionable — don't change those. Focus on the token acquisition error paths.

2. **Enhance `src/client.js` error handling** — modify the two `throw new Error` lines:

   In `request()` method (line ~29, handles Graph API):
   ```javascript
   if (!res.ok) {
     const errBody = await res.text();
     if (res.status === 401) throw new Error("Authentication expired or revoked. Use the 'disconnect' tool to clear tokens and reconnect.");
     if (res.status === 403) throw new Error(`Access denied: Your account may not have permission for this operation. (${errBody})`);
     if (res.status === 404) throw new Error(`Resource not found: Verify the site URL and resource ID are correct. (${errBody})`);
     throw new Error(`SharePoint API error ${res.status}: ${errBody}`);
   }
   ```

   In `spRequest()` method (line ~63, handles SP REST API):
   ```javascript
   if (!res.ok) {
     const errBody = await res.text();
     if (res.status === 401) throw new Error("Authentication expired or revoked. Use the 'disconnect' tool to clear tokens and reconnect.");
     if (res.status === 403) throw new Error(`Access denied: Your account may not have permission for this operation. (${errBody})`);
     if (res.status === 404) throw new Error(`Resource not found: Verify the site URL and resource ID are correct. (${errBody})`);
     throw new Error(`SharePoint API error ${res.status}: ${errBody}`);
   }
   ```

3. **Add error message tests to `tests/client.test.js`** — add tests that:
   - Mock `fetch` to return a 401 response → verify error message contains "Authentication expired"
   - Mock `fetch` to return a 403 response → verify error message contains "Access denied"
   - Mock `fetch` to return a 404 response → verify error message contains "Resource not found"
   - Mock `fetch` to return a 500 response → verify error message starts with "SharePoint API error"
   - Test both `request()` (Graph) and `spRequest()` (SP REST) paths if feasible

   Look at the existing test patterns in `tests/client.test.js` for how fetch is mocked (likely injectable fetchFn or similar pattern matching auth.test.js).

4. **Add auth error tests to `tests/auth.test.js`** — add tests that:
   - Verify Conditional Access error code produces "Conditional Access" guidance
   - Verify generic AADSTS error produces "disconnect" guidance
   - Look at existing auth test structure first — if MSAL is deeply mocked, test at whatever level is practical without over-engineering

5. **Run full regression** — `node --test tests/*.test.js` — all tests pass (existing 56 + new error message tests).

6. **Verify tool count invariant** — `grep -c 'server.tool(' src/tools.js` still equals 25.

## Must-Haves

- 401/403/404 errors in client.js produce distinct actionable messages
- At least Conditional Access AADSTS codes produce specific guidance in auth.js
- New error message tests exist and pass
- All 56 existing tests still pass (no regression)
- No changes to tool registrations (tool count invariant)

## Inputs

- `src/auth.js` — error paths at lines 44 (tenant discovery), 141 (silent acquisition), and in the device code flow. Tenant discovery errors are already good — focus on token acquisition.
- `src/client.js` — two throw sites: line ~29 (`API ${res.status}: ${errBody}`) and line ~63 (`SP REST ${res.status}: ${errBody}`)
- `tests/auth.test.js` — existing auth tests (part of 56 total)
- `tests/client.test.js` — existing client tests (part of 56 total)
- Test pattern: injectable `fetchFn` for network mocking, `node:test` + `node:assert/strict`

## Expected Output

- `src/auth.js` — enhanced error messages for AADSTS codes in token acquisition
- `src/client.js` — enhanced error messages for 401/403/404 HTTP responses
- `tests/auth.test.js` — new tests for auth error guidance messages
- `tests/client.test.js` — new tests for client error guidance messages
- All tests pass (56 existing + new)

## Observability Impact

- Error messages now include actionable guidance — this improves the diagnostic surface for end users
- Original error details are preserved (in parentheses or appended) for debugging by developers
- AADSTS codes in auth errors enable programmatic detection of specific failure modes

## Verification

```bash
node --test tests/*.test.js
grep -c 'server.tool(' src/tools.js  # must be 25
```
