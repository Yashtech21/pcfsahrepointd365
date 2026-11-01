# API Specifications

There are two layers: **our Document API** (called by the PCF control) and the **Microsoft Graph** calls the
backend makes. The control never calls Graph.

---

## Part A: Document API (Azure Function, `backend/src/functions/documents.ts`)

Base URL: the `apiBaseUrl` control property, e.g. `https://<app>.azurewebsites.net/api`.

**Every request** needs `Authorization: Bearer <Entra access token>`, with audience in `API_AUDIENCE` and the
tenant `GRAPH_TENANT_ID`. Missing or invalid token → `401`. Wrong tenant or missing role → `403`.

All errors return JSON `{ "error": "<message>" }`.

### DocumentDto (response item)

```json
{
  "id": "01ABCDEF...",            // Graph driveItem id
  "name": "Form_U4_Verification.pdf",
  "category": "FinCEN_U4",        // DocCategory value ("" if untagged)
  "sourceEntity": "ce_advisoronboarding", // "contact" when SourceEntity is empty
  "sourceRecordId": "e5f6…",      // lowercase GUID ("" if none)
  "contactId": "a1b2…",           // lowercase GUID
  "uploadDate": "2022-03-11T10:00:00Z", // driveItem createdDateTime
  "webUrl": "https://contoso.sharepoint.com/…",
  "size": 402000,
  "fileType": "pdf"
}
```

The client type is `IDocumentItem` in `pcf/.../types.ts` (it doesn't use `contactId`).

### Endpoints

| Method & route | Request | Success | Errors |
|---|---|---|---|
| `GET /documents?contactId={guid}` | — | `200 DocumentDto[]` (files only, all pages) | `400` bad GUID |
| `POST /documents/upload?fileName=&category=&contactId=&contactName=&sourceEntity=&sourceRecordId=` | **Body: raw file bytes**, `Content-Type` = file MIME | `201 DocumentDto` | `400` validation, `413` > `MAX_UPLOAD_MB`, `415` file rejected (see "File safety"), `502` stored size mismatch (upload discarded) |
| `POST /documents/{itemId}/preview` | no body | `200 { "url": "…" }` **or** `200 { "url": "…", "postParameters": "a=1&b=2" }` (POST the form-encoded parameters to `url` inside the iframe) | `400` not a file / no preview, `404` |
| `GET /documents/{itemId}/download` | — | `200 { "url": "<pre-authenticated downloadUrl>" }` | `400` not a file, `404` |
| `DELETE /documents/{itemId}` | header `X-Confirm-Delete: <encodeURIComponent(fileName)>` | `204` | `400` missing header / folder, `404`, `409` name doesn't match the item |

Shared error mapping (`handler()` in `documents.ts`): Graph `404` → `404`; Graph `403` → `403`; any timeout
(Graph call or upload chunk) → `504`; Graph `413`/`507` (quota) → `507`; other Graph or network errors → `502`;
anything unexpected → `500`. Every URL returned to the browser (preview, download) has passed `assertTrustedUrl`
(https and a Microsoft/SharePoint host).

### Validation rules (server)

- `contactId`, `sourceRecordId`: GUID (braces and case are normalised).
- `sourceEntity`: `^[a-z][a-z0-9_]{0,63}$`.
- `sourceEntity` must have an **active** Document Manager Configuration row → otherwise `400 Document management is
  not configured for table "x"`.
- `category` must be an **active** category in that row's Allowed Categories → otherwise `400 Category "x" is not
  allowed on <table>`. Its folder comes from the category's Folder Path (`safeFolderPath`).
- Size limit = min(`MAX_UPLOAD_MB`, the row's Max Upload) → `413 File is too large (limit N MB for <table>)`.
- `fileName`: required. It is sanitised by `sanitizeName` (`naming.ts`): illegal characters, control characters,
  the `~$` prefix, `_vti_`, leading or trailing dots/spaces and reserved device names are handled, and the
  extension is kept when shortening.

### File safety (`backend/src/lib/fileRules.ts`, checked before anything is written)

- Blocked extensions: `DEFAULT_BLOCKED` (executables, scripts, disk images, macro-enabled Office). The **final**
  extension is checked, so `invoice.pdf.exe` is blocked. Override with the `BLOCKED_EXTENSIONS` app setting.
- Executable content (PE/ELF/Mach-O headers, `#!` scripts) is rejected under **any** name.
- Content must match the extension for the types in `SIGNATURES` (pdf, Office, OpenDocument, images, rtf, zip).
  An `.exe` renamed to `.pdf` is rejected.
- After upload, the stored size must equal the received size. Otherwise the file is deleted and `502` returned.
- `itemId`: `^[A-Za-z0-9!_\-]+$`.

### Why raw bytes and not base64 JSON

The original spec suggested `{ …, "fileContentBase64": "…" }`. We send the raw body with metadata in the query
string instead: it is about 33% smaller on the wire, avoids holding a second encoded copy in memory, and needs no
JSON parsing of large payloads. **Do not reintroduce base64 uploads.**

---

## Part B: Microsoft Graph calls (backend only)

Source of truth: `backend/src/lib/sharepoint.ts` (requests), `graph.ts` (transport) and `uploadSession.ts` (large
files). This section describes **behaviour and contracts**. For exact query strings, timeouts and sizes, read the
named constants in code rather than copying them here, so the doc can't drift.

### Transport rules (`graph()` in `graph.ts`)

- App-only token, fetched per attempt. Every attempt has a timeout (`REQUEST_TIMEOUT_MS`, app setting
  `GRAPH_TIMEOUT_MS`).
- Retries follow `retryDecision()`. `429`/`503` are always retried after `Retry-After` (seconds or HTTP-date,
  via `retryDelayMs`). Timeouts, network errors and `500/502/504` are retried **only for idempotent requests**.
- Idempotent by default: GET, HEAD, DELETE, PATCH. POST/PUT are not, unless the caller passes `idempotent: true`
  (done for preview, folder creation with `conflictBehavior=fail`, and `createUploadSession`).
  **The small-file upload PUT is deliberately not idempotent.** With `conflictBehavior=rename`, repeating it after a
  timeout could create a duplicate.
- Failures always surface as `GraphError`.

### Operations

| Purpose | Graph call(s) | Notes |
|---|---|---|
| Resolve site, library, drive | `GET /sites/{host}:{path}`, `GET …/lists`, `GET …/lists/{id}/drive` | `getTarget()`, cached per worker |
| Contact folder | `GET root:/{name}` → `GET root/search(q='{guid}')` → `POST root/children` (`conflictBehavior: fail`) | `ensureContactFolder()`. Matches on the `_{guid}` suffix. Cache bounded by `FOLDER_CACHE_MAX` |
| Upload ≤ `SMALL_UPLOAD_LIMIT` | `PUT root:/{path}:/content?@microsoft.graph.conflictBehavior=rename` | Intermediate folders are created by Graph |
| Upload > `SMALL_UPLOAD_LIMIT` | `POST root:/{path}:/createUploadSession`, then `uploadToSession()` | See "Large-file uploads" |
| Tag metadata | `PATCH /items/{id}/listItem/fields` (the four columns) | On failure the file is deleted, then the error is rethrown |
| List by contact | `GET /lists/{id}/items?$expand=fields(…),driveItem(…)&$filter=fields/ContactGuid eq '{guid}'` | Follows `@odata.nextLink`. Files only. Columns in `FIELDS` |
| Preview | `GET /items/{id}?$select=id,name,file`, then `POST /items/{id}/preview` | Office types (`isOfficeFile`) send `{ viewer: "office", chromeless: true }`, others `{}`. Returns `getUrl`, or `postUrl` + `postParameters` |
| Download URL | `GET /items/{id}` → `@microsoft.graph.downloadUrl` | Short-lived and pre-authenticated. Files only |
| Delete | `GET /items/{id}?$select=id,name,file,folder`, then `DELETE /items/{id}` | Must be a file, and the name must equal `X-Confirm-Delete` |

### Large-file uploads (`uploadToSession()` in `uploadSession.ts`)

- Chunks of `CHUNK_SIZE` (must stay a multiple of 320 KiB), each with `Content-Range` and its own timeout
  (`CHUNK_TIMEOUT_MS`). The pre-authenticated `uploadUrl` **never** receives the bearer token.
- `202` → continue at `nextExpectedRanges`. `200/201` → the response body is the created item.
- `429/503` → wait for `Retry-After` and resend the same range.
- Timeout, network error, `5xx` or `416` → `GET uploadUrl` to read `nextExpectedRanges` and **resume from there**,
  so a chunk that landed is never sent twice.
- `404` (session expired) or any other `4xx` → stop.
- After `MAX_FAILURES` consecutive failures, or on any stop: `DELETE uploadUrl` (with `CONTROL_TIMEOUT_MS`), then
  throw `GraphError` (`504` for timeouts).
- Covered by `backend/test/uploadSession.test.ts`.

### Permissions

App-only **`Sites.Selected`** with a **Write** grant on `/sites/AdvisorManagement` (see
`sharepoint/Setup-ContactsLibrary.ps1`). Don't request `Sites.ReadWrite.All` or `Files.ReadWrite.All`.

---

## Part C: Dataverse

### Control (as the signed-in user, read-only, `services/SettingsService.ts`, cached per page load)

| Purpose | Call |
|---|---|
| Entity display name | `context.utils.getEntityMetadata(logicalName)` → `DisplayName` |
| User UPN (MSAL login hint) | `context.webAPI.retrieveRecord("systemuser", userId, "?$select=domainname")` |
| Connection settings | `retrieveMultipleRecords("environmentvariabledefinition", …$filter=schemaname eq 'ce_DocManager…'&$expand=environmentvariabledefinition_environmentvariablevalue)`. The current value wins over the default; non-empty form properties override both |
| Table configuration | `retrieveMultipleRecords("ce_documentmanagerconfig", …$filter=ce_tablelogicalname eq '<table>' and statecode eq 0&$expand=ce_documentmanagerconfig_documentcategory(…))` |
| Category catalog | `retrieveMultipleRecords("ce_documentcategory", "?$select=ce_name,ce_value,ce_color,ce_sortorder,statecode")` |

### Backend (as its application user, `backend/src/lib/tableConfig.ts`, `dataverse()` in `graph.ts`)

| Purpose | Call |
|---|---|
| Table configuration + allowed categories | `GET {DATAVERSE_URL}/api/data/v9.2/ce_documentmanagerconfigs?$filter=ce_tablelogicalname eq '<table>' and statecode eq 0&$expand=ce_documentmanagerconfig_documentcategory($select=ce_name,ce_value,ce_folderpath,ce_color,ce_sortorder,statecode)` |

Token: client credentials for `{DATAVERSE_URL}/.default` (same app registration as Graph). Same timeout and retry
rules as Graph. Errors surface as `Dataverse error: …`.

There is no Dataverse Custom API or Power Automate flow in this design (that was "Option B" in the original spec and
was not built).
