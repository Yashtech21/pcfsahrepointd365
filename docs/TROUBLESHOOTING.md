# Troubleshooting & Known Quirks

Each entry: symptom → cause → fix. Entries marked **(hit during development)** actually happened in this repo.
Don't "fix" them by undoing the documented workaround.

## Build (`pcf/`)

1. **`Error: Could not find config file.` right after "Running ESLint…"** *(hit during development)*
   - Cause: `pcf-scripts` 1.5x runs ESLint 9 and the repo has no flat config.
   - Fix: `pcfconfig.json` sets `"skipBuildLinting": true`. Use `npx tsc --noEmit` for type checks. If you want linting,
     add an `eslint.config.mjs` and remove the flag in the same change.

2. **`Unsupported 'react' version '18.3.1' specified platform-library`** *(hit during development)*
   - Cause: React 18 / Fluent 9.68 platform libraries are behind the `pcfReactPlatformLibraries` preview flag.
   - Fix: stay on `React 16.14.0` + `Fluent 9.46.2` in the manifest, with matching `devDependencies`. Don't enable
     the preview flag for production.

3. **`Can't resolve 'react/jsx-runtime' … @griffel/react … fully specified`** *(hit during development)*
   - Cause: `@fluentui/react-icons` pulls its own Griffel ESM build, which imports `react/jsx-runtime` without an
     extension. React 16 has no `exports` map, so webpack's strict ESM resolution fails.
   - Fix: don't install `@fluentui/react-icons`. Add icons to `components/Icons.tsx`.

4. **`TS2322: Type 'string' is not assignable to type 'undefined'` inside `makeStyles`** *(hit during development)*
   - Cause: Griffel rejects CSS shorthands (`borderColor`, `border`, multi-value `padding`, …) in its types.
   - Fix: use `shorthands.borderColor(...)` (etc.) or the longhand properties.

5. **Using a Fluent component that compiles but is `undefined` at runtime in D365**
   - Cause: the dev dependency is newer than the platform library (9.46.2) that D365 actually loads.
   - Fix: keep `@fluentui/react-components` pinned to `9.46.2`.

6. **Bundle suddenly > 1 MB**
   - Cause: React, Fluent or Griffel got bundled (wrong import path or a new dependency).
   - Fix: `grep -c griffel out/controls/CategorizedDocManager/bundle.js` should be 0. Remove the offending import.

## Build (`backend/`)

7. **`Type 'Buffer' is not assignable to type 'BodyInit'`** *(hit during development)*
   - Cause: Node's `fetch` types don't accept `Buffer`.
   - Fix: pass `new Uint8Array(buffer)` (or `new Uint8Array(buffer.subarray(a, b))` for chunks).

8. **`npm test` → `Cannot find module dist/test`** *(hit during development)*
   - Fix: the script is `node --test dist/test/*.test.js`. Keep the glob.

## Packaging / deployment

9. **`pac pcf push` fails with an authentication error**
   - Fix: `pac auth create --url https://<org>.crm.dynamics.com`, then `pac auth select --index <n>`.

10. **`pac pcf push` rejects the version / the control doesn't update**
    - Fix: bump `version` in `ControlManifest.Input.xml`, or hard-refresh the form (Ctrl+F5).

11. **`dotnet build` of the solution can't find the PCF project**
    - Fix: run `pac solution add-reference --path ../pcf` from the solution folder. The `.pcfproj` is
      `pcf/CategorizedDocManager.pcfproj`.

## Runtime: auth

12. **Sign-in popup every session / popup blocked**
    - Cause: `ssoSilent` failed: third-party cookies blocked, or the SPA redirect URI isn't registered for this org URL.
    - Fix: add `https://<org>.crm.dynamics.com` as an SPA redirect URI on the client app registration. Allow popups for the org.

13. **`401 Invalid or expired token`**
    - Cause: the token audience isn't in `API_AUDIENCE`, or the control's `apiScope` points at a different app.
    - Fix: set `API_AUDIENCE=api://<backend-client-id>,<backend-client-id>` and make the scope
      `api://<backend-client-id>/user_impersonation`.

14. **`403 Caller lacks the required role`**: assign the `DocManager.User` app role, or clear `REQUIRED_ROLE`.

15. **Browser console: CORS error on `/api/documents`**
    - Fix: add the org URL to Function App → CORS. The request uses the `Authorization` and `X-Confirm-Delete`
      headers, so any proxy or APIM in front must allow both.

## Runtime: SharePoint / Graph

16. **`SharePoint error: Access denied` / Graph 403**
    - Cause: missing `Sites.Selected` grant, or admin consent not given.
    - Fix: re-run `sharepoint/Setup-ContactsLibrary.ps1` and grant admin consent for `Sites.Selected`.

17. **Preview pane shows an error, or the iframe is blank or blocked**
    - Cause: the file type has no Graph preview, or the iframe `src` isn't the exact `getUrl`.
    - Fix: always use the `getUrl` returned by `/preview` unmodified. Office files use `viewer: "office", chromeless: true`;
      PDFs and images must send `{}` (the Office viewer can't render them). Users can still use "Open in SharePoint".

18. **Large upload fails partway**
    - Cause: chunk size isn't a multiple of 320 KiB, a chunk took longer than 120 s, or the bearer token was sent to `uploadUrl`.
    - Fix: keep `CHUNK_SIZE = 16 * 320 * 1024` (5 MiB) and send no `Authorization` header on `uploadUrl`.

19. **"did not respond within N seconds" banner**
    - Cause: network stall, a slow token refresh, or Graph throttling that went past the timeout.
    - Fix: use **Retry**. If it happens often on big contacts, raise `GRAPH_TIMEOUT_MS` (backend) or
      `DEFAULT_TIMEOUT_MS` (`DocumentApi.ts`). Don't remove the timeouts.

20. **A document exists in SharePoint but not in the control**
    - Cause: it has no `ContactGuid` (uploaded outside the control) or isn't a file.
    - Fix: fill in `ContactGuid` (and ideally the other three columns) in SharePoint.

21. **`409 Delete was not confirmed for this file`**
    - Cause: the file was renamed in SharePoint after the list loaded.
    - Fix: refresh and delete again.

22. **Contact renamed: is a new folder created?**
    - No. The folder is matched by its `_{guid}` suffix. A brand-new folder only appears if search hasn't indexed the
      old one yet, which is rare and only right after creation.
