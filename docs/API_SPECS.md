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
| `POST /documents/upload?fileName=&category=&contactId=&contactName=&sourceEntity=&sourceRecordId=` | **Body: raw file bytes**, `Content-Type` = file MIME | `201 DocumentDto` | `400` validation, `413` > `MAX_UPLOAD_MB` (default 100) |
| `POST /documents/{itemId}/preview` | no body | `200 { "url": "<embed getUrl>" }` | `400` not a file / no preview, `404` |
| `GET /documents/{itemId}/download` | — | `200 { "url": "<pre-authenticated downloadUrl>" }` | `400` not a file, `404` |
| `DELETE /documents/{itemId}` | header `X-Confirm-Delete: <encodeURIComponent(fileName)>` | `204` | `400` missing header / folder, `404`, `409` name doesn't match the item |

Shared error mapping: Graph `404` → `404`; Graph `403` → `403`; Graph or our own timeout → `504`; other Graph
errors → `502`; anything unexpected → `500`.

### Validation rules (server)

- `contactId`, `sourceRecordId`: GUID (braces and case are normalised).
- `sourceEntity`: `^[a-z][a-z0-9_]{0,63}$`.
- `category`: must be a key of `CATEGORY_FOLDERS`.
- `fileName`: required. It is sanitised (`"*:<>?/\|#%` → `_`, control characters and trailing dots/spaces removed,
  max 200 chars).
- `itemId`: `^[A-Za-z0-9!_\-]+$`.

### Why raw bytes and not base64 JSON

The original spec suggested `{ …, "fileContentBase64": "…" }`. We send the raw body with metadata in the query
string instead: it is about 33% smaller on the wire, avoids holding a second encoded copy in memory, and needs no
JSON parsing of large payloads. **Do not reintroduce base64 uploads.**

---

## Part B: Microsoft Graph calls (backend only, `backend/src/lib/sharepoint.ts`)

All calls go through `graph()` (app-only token, 30 s timeout via `GRAPH_TIMEOUT_MS`, retries on 429/503/504
honouring `Retry-After` given in seconds or as an HTTP-date, capped at 30 s, with an exponential fallback).

| Purpose | Request |
|---|---|
| Resolve site | `GET /sites/{hostname}:{sitePath}` |
| Resolve library | `GET /sites/{siteId}/lists?$select=id,name,displayName` → match `SP_LIBRARY_NAME` |
| Resolve drive | `GET /sites/{siteId}/lists/{listId}/drive?$select=id` |
| Contact folder exists? | `GET /drives/{driveId}/root:/{encoded folder}?$select=id` |
| Find renamed folder | `GET /drives/{driveId}/root/search(q='{contactGuid}')?$select=id,name,folder,parentReference&$top=50` |
| Create contact folder | `POST /drives/{driveId}/root/children` `{ "name": "...", "folder": {}, "@microsoft.graph.conflictBehavior": "fail" }` (a 409 means another request created it) |
| Upload ≤ 4 MB | `PUT /drives/{driveId}/root:/{path}:/content?@microsoft.graph.conflictBehavior=rename` (body = bytes) |
| Upload > 4 MB | `POST /drives/{driveId}/root:/{path}:/createUploadSession` `{ "item": { "@microsoft.graph.conflictBehavior": "rename" } }`, then `PUT {uploadUrl}` chunks of **5 MiB = 16 × 320 KiB** with `Content-Range: bytes {start}-{end}/{total}`. **No** `Authorization` header on `uploadUrl`. 120 s per chunk. On failure the session is cancelled with `DELETE {uploadUrl}` |
| Tag metadata | `PATCH /drives/{driveId}/items/{itemId}/listItem/fields` `{ "DocCategory", "SourceEntity", "SourceRecordId", "ContactGuid" }`. On failure → `DELETE` the item, then rethrow |
| List by contact | `GET /sites/{siteId}/lists/{listId}/items?$expand=fields($select=DocCategory,SourceEntity,SourceRecordId,ContactGuid),driveItem($select=id,name,webUrl,size,createdDateTime,file)&$filter=fields/ContactGuid eq '{guid}'&$top=200` + header `Prefer: HonorNonIndexedQueriesWarningMayFailRandomly`. Follow `@odata.nextLink`. Keep only items with `driveItem.file` |
| Get item | `GET /drives/{driveId}/items/{id}?$expand=listItem($expand=fields($select=…))` |
| Preview | `POST /drives/{driveId}/items/{id}/preview`. Body: **Office files** (`doc[x] xls[x] ppt[x] odt ods odp rtf`) `{ "viewer": "office", "chromeless": true }`; **everything else** (PDF, images) `{}`. Use the response's `getUrl` as the iframe `src`. If there is no `getUrl`, return 400 |
| Download URL | `GET /drives/{driveId}/items/{id}` → `@microsoft.graph.downloadUrl` (short-lived, pre-authenticated) |
| Delete | `GET /drives/{driveId}/items/{id}?$select=id,name,file,folder` (must be a file and the name must match), then `DELETE /drives/{driveId}/items/{id}` |

### Permissions

App-only **`Sites.Selected`** with a **Write** grant on `/sites/AdvisorManagement` (see
`sharepoint/Setup-ContactsLibrary.ps1`). Don't request `Sites.ReadWrite.All` or `Files.ReadWrite.All`.

---

## Part C: Dataverse (client only, read-only)

| Purpose | Call |
|---|---|
| Entity display name | `context.utils.getEntityMetadata(logicalName)` → `DisplayName` |
| User UPN (MSAL login hint) | `context.webAPI.retrieveRecord("systemuser", userId, "?$select=domainname")` |

There is no Dataverse Custom API or Power Automate flow in this design (that was "Option B" in the original spec and
was not built).
