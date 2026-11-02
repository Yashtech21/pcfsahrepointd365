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

17a. **"Graph returned an untrusted URL host"**
    - Cause: a preview or download URL wasn't https on a Microsoft/SharePoint host (`assertTrustedUrl`).
    - Fix: if Microsoft introduces a new legitimate host (sovereign clouds, new CDN), add its suffix to the list in
      `assertTrustedUrl`, with a test in `safety.test.ts`.

17. **Preview pane shows an error, or the iframe is blank or blocked**
    - Cause: the file type has no Graph preview, or the iframe `src` isn't the exact `getUrl`.
    - Fix: always use the `getUrl` returned by `/preview` unmodified. Office files use `viewer: "office", chromeless: true`;
      PDFs and images must send `{}` (the Office viewer can't render them). Users can still use "Open in SharePoint".

18. **Large upload fails partway**
    - `uploadToSession` already retries throttling and resumes after timeouts from `nextExpectedRanges`. A final
      failure means `MAX_FAILURES` was hit, the session expired (404), or SharePoint returned a non-retryable 4xx.
    - Check: `CHUNK_SIZE` is still a multiple of 320 KiB, no `Authorization` header goes to `uploadUrl`, and the
      Function's own timeout (`functionTimeout` in `host.json`, 5 min default on Consumption) covers the largest
      allowed file.

19. **"did not respond within N seconds" banner**
    - Cause: network stall, a slow token refresh, or Graph throttling that went past the timeout.
    - Fix: use **Retry**. If it happens often on big contacts, raise `GRAPH_TIMEOUT_MS` (backend) or
      `DEFAULT_TIMEOUT_MS` (`DocumentApi.ts`). Don't remove the timeouts.

20a. **`415` "Files of type .x are not allowed" / "content does not match its extension" / "Executable files…"**
    - Cause: `fileRules.ts` rejected the file (blocked type, disguised file, or executable content).
    - Fix: this is intended. If a legitimate type is blocked (e.g. macro-enabled `.xlsm` your business needs), set
      `BLOCKED_EXTENSIONS` on the Function App to the full list you want. Setting it replaces the default list.

20b. **Upload timed out in the browser, but the file shows up anyway**
    - Expected: the server finished after the client gave up. The control detects this
      (`landedDespiteTimeout`) and doesn't leave the file staged, so it isn't uploaded twice.

20. **A document exists in SharePoint but not in the control**
    - Cause: it has no `ContactGuid` (uploaded outside the control) or isn't a file.
    - Fix: fill in `ContactGuid` (and ideally the other three columns) in SharePoint.

21. **`409 Delete was not confirmed for this file`**
    - Cause: the file was renamed in SharePoint after the list loaded.
    - Fix: refresh and delete again.

22. **Contact renamed: is a new folder created?**
    - No. The folder is matched by its `_{guid}` suffix. A brand-new folder only appears if search hasn't indexed the
      old one yet, which is rare and only right after creation.

## Configuration (Dataverse)

23. **"The Document Manager isn't connected yet. Set these environment variables: …"**
    - Cause: the listed `ce_DocManager*` environment variables have no value (and no override on the form).
    - Fix: `node dataverse/provision.mjs set-env --org … --api-base-url … --tenant-id … --client-id … --api-scope …`,
      or **Solutions → DocumentManager → Environment variables → New value**.

24. **"Not set up for this table"**
    - Cause: no **active** Document Manager Configuration row whose Table Logical Name equals the form's table.
    - Fix: add or activate the row (docs/CONFIGURATION.md §3).

25. **"Could not read Document Manager configuration / document categories"**
    - Cause: the tables don't exist in this environment, or the user lacks the **Document Manager Reader** role.
    - Fix: run `provision.mjs all` for this org; assign the role (directly or via a team).

26. **Warning "This control is bound to "x", but the configuration … says the contact lookup is "y""**
    - Cause: the form binds the control to a different lookup than the row's **Contact Lookup Column**.
    - Fix: correct whichever is wrong. Uploads check the record's contact through the **configured** column, so a
      wrong value makes every upload fail with a 409.

27. **New category doesn't appear**
    - Control: it's read once per page load, so reload the form. Check that the category is **active** and added to
      the table's **Allowed Categories**.
    - Backend: up to `CONFIG_CACHE_SECONDS` (default 5 min) before uploads with the new category are accepted.

28. **Upload fails with `Dataverse error: …` or `Could not reach Dataverse`**
    - Cause: the backend can't read configuration: `DATAVERSE_URL` is wrong or missing, the application user is
      missing (`provision.mjs app-user`), or it lacks the **Document Manager Reader** role.

29. **`provision.mjs`: "Could not get a token with Azure CLI"**
    - Fix: `az login` with your D365 admin account (use `--tenant <tenant-id>` if you have several), or set
      `DATAVERSE_TOKEN` to a bearer token for the org.

30. **`provision.mjs role`: "Privilege prvReadce_DocumentCategory not found"**
    - Cause: the tables weren't created or published yet. Run `schema` (it publishes), then `role` again.

## Per-record security

31. **`403 You don't have access to this contact` / `… this record`**
    - Expected when the user's D365 roles don't let them read it. Check in D365 whether the user can open that
      contact or record. If they can, see 34.

32. **`403 You can view this record but don't have permission to change its documents`**
    - Upload and delete need **Write** on the source record (or on the contact for contact-level documents). Grant it
      through their role or by sharing, or set `CHANGE_ACCESS=read` on the Function App.

33. **`409 This <table> is linked to a different contact` / `has no contact selected`**
    - The record's contact (the configured Contact Lookup Column) differs from the one on the form, usually because
      the form wasn't saved after the contact changed. Save, refresh and retry. If it always happens, the
      configuration row's Contact Lookup Column is wrong.

34. **Every request fails with `Dataverse error: Could not get a Dataverse token for the user`**
    - The backend app lacks **Dynamics CRM → user_impersonation (Delegated)** or its admin consent (SETUP 1a.2), or the
      client secret is wrong or expired.

35. **A document disappeared from "All Contact Documents" for some users**
    - Its source record isn't readable by them (expected), **or** the record was deleted: Dataverse can't tell those
      apart, so such documents are hidden from everyone. They're still in SharePoint. Re-tag them to the contact
      (`SourceEntity` = `contact`, clear `SourceRecordId`) to make them visible at contact level.

36. **`403 This document isn't linked to a contact`**
    - The file has no `ContactGuid` (e.g. uploaded straight into SharePoint). Fill in its four columns.

## Copy link

37. **Someone opens a copied link and gets "Access denied" in SharePoint**
    - Expected: the link is the file's plain SharePoint URL and grants nothing. The recipient needs Read on the
      AdvisorManagement site or the file. Copy link deliberately doesn't create sharing links.

38. **"Your browser didn't allow automatic copying" dialog**
    - The browser blocked the Clipboard API (policy, permissions or an embedded context). Press Ctrl+C (the link is
      already selected) or use **Try again**.

39. **Copy link button is disabled ("Link not available")**
    - The backend dropped the file's link because it wasn't an https Microsoft/SharePoint URL (`assertTrustedUrl`).

