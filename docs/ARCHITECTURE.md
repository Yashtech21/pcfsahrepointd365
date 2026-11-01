# Architecture Specification

## 1. Components and boundaries

| Component | Location | Responsibility | Must NOT |
|---|---|---|---|
| PCF control | `pcf/CategorizedDocManager/` | UI, reading form context, user sign-in (MSAL), calling our API | Call Graph/SharePoint, hold secrets, decide folder paths |
| Document API | `backend/` (Azure Function) | Auth, validation, folder resolution, uploads, metadata, list/preview/download/delete | Trust client-provided folder paths or the Easy Auth header |
| SharePoint | `/sites/AdvisorManagement`, library `Contacts` | Storage + metadata | — |
| Dataverse | — | Hosts the form; the control reads entity metadata and the user's UPN only | Store document data (no SharePoint Document Locations are used) |

```
┌──────────────── Dynamics 365 model-driven form (Opportunity / Advisor Onboarding) ────────────────┐
│ CategorizedDocManager (PCF virtual control, bound to the Contact lookup)                          │
│   index.ts ─ caches host/categories/api ─► DocumentManager.tsx                                    │
│      UploadBar · TabList · DocumentCard grid · PreviewPane · DeleteDialog · SkeletonCards         │
│   services/AuthService (MSAL ssoSilent) ─► services/DocumentApi (fetchWithTimeout)                │
└──────────────────────────────────────┬────────────────────────────────────────────────────────────┘
                                       │ HTTPS, Authorization: Bearer <user token for api://…/user_impersonation>
┌──────────────────────────────────────▼────────────────────────────────────────────────────────────┐
│ Azure Function  /api/documents/*   (functions/documents.ts)                                        │
│   lib/auth.ts      verify JWT (JWKS, issuer, audience, tenant, optional role)                       │
│   lib/sharepoint.ts folder resolution, upload (simple / chunked), metadata, list, preview, delete   │
│   lib/graph.ts     app-only token (client secret), timeouts, 429/503/504 retry with Retry-After     │
└──────────────────────────────────────┬────────────────────────────────────────────────────────────┘
                                       │ Microsoft Graph v1.0, app-only, Sites.Selected (Write on one site)
                              SharePoint: /sites/AdvisorManagement/Contacts
```

## 2. Information architecture

Documents are **not** stored in per-entity folders (the out-of-the-box `/opportunity/<guid>` model is deliberately
not used). Everything for an advisor lives under the parent Contact:

```
/sites/AdvisorManagement/Contacts/{Contact_FullName}_{Contact_GUID}/
├── Offer Letters/
├── Commissions/
├── Onboarding Documents/
│   ├── Payslips/
│   ├── Resignation Letters/
│   └── FinCEN_U4/
└── Miscellaneous/
```

- Folder name = `sanitizeName(fullName)` (max 120 chars) + `_` + lowercase GUID (`backend/src/lib/naming.ts`).
- **The GUID suffix identifies the folder, not the name.** If the contact is renamed, `ensureContactFolder`
  finds the existing folder by its `_{guid}` suffix (Graph search) and keeps using it. It does not rename folders.
- Folders are created lazily. The contact root is created explicitly, and category sub-folders are created
  implicitly by Graph path-based uploads.
- Conflicting file names are auto-renamed by Graph (`conflictBehavior=rename`), never overwritten.
- Files are validated before anything is written (`fileRules.ts`: blocked types, executable content,
  content/extension match). After upload, the stored size must equal the received size, otherwise the file is deleted.

### Categories, folders and tables come from Dataverse

Categories (`ce_documentcategory`: value, label, folder path, colour, order) and per-table configuration
(`ce_documentmanagerconfig`: table logical name, **Contact lookup column**, max upload, allowed categories) are
Dataverse data. See [CONFIGURATION.md](CONFIGURATION.md).
- **Backend** (`backend/src/lib/tableConfig.ts`) reads them as its application user, cached per table for
  `CONFIG_CACHE_SECONDS`, and is the enforcement point: the table must be configured and active, the category must be
  active and allowed on that table, and the folder path is sanitised (`safeFolderPath`).
- **Control** (`services/SettingsService.ts`) reads them as the user, for the dropdown, labels, colours and the
  client-side size check. Inactive categories still label existing documents.
- The starting data is `dataverse/seed.json`, which is also the backend's local-dev config (`DOC_CONFIG_FILE`).

## 3. SharePoint columns (library `Contacts`)

| Internal name | Type | Indexed | Value written on upload |
|---|---|---|---|
| `DocCategory` | Single line of text | yes | category **Value** from Dataverse |
| `SourceEntity` | Single line of text | no | logical name of the hosting entity, e.g. `opportunity`, `ce_advisoronboarding` |
| `SourceRecordId` | Single line of text | yes | lowercase GUID of the hosting record |
| `ContactGuid` | Single line of text | yes | lowercase GUID of the Contact |

- Listing filters on `ContactGuid`. The index is what keeps it working beyond the 5,000-item list view threshold.
- A file with no `ContactGuid` (e.g. uploaded directly in SharePoint) is **invisible** to the control.
- If writing metadata fails after an upload, the backend **deletes the uploaded file** before returning the error,
  so there are never untagged files created by the control.
- When `SourceEntity` is empty the UI treats the file as `contact` and shows "Contact (Root)".

## 4. Dual-scope view rules

Both tabs use **one** list call: `GET /api/documents?contactId=…`. Tabs only filter in memory.

- **"This Record's Documents ({entity display name})"**: `doc.sourceRecordId === host.recordId`. Disabled
  while the record is unsaved (no id). When the record gets an id after its first save, the view switches to this tab.
- **"All Contact Documents ({contact name})"**: every file tagged with the contact, from any entity.
- Upload is disabled with a message when there is no contact selected or the record isn't saved yet.

## 5. Host context resolution (PCF)

| Value | Source |
|---|---|
| Entity logical name / record id | `context.mode.contextInfo.entityTypeName` / `.entityId` (normalised to lowercase, no braces) |
| Entity display name | `context.utils.getEntityMetadata(name).DisplayName`, cached per entity and loaded into React state |
| Contact id / name | the bound `contactLookupField` (`Lookup.Simple`) `raw[0].id` / `.name` |
| User UPN (MSAL login hint) | `context.webAPI.retrieveRecord("systemuser", userId, "?$select=domainname")`, once in `init` |

## 6. Security model

1. **User → API:** MSAL obtains a token for `api://<backend-app>/user_impersonation` (SPA app registration).
2. **API checks the caller:** `auth.ts` verifies the signature against the tenant JWKS, issuer (v1 or v2), audience
   (`API_AUDIENCE`), expiry, `tid`, and optionally the `REQUIRED_ROLE` app role. App Service Authentication is
   optional defense-in-depth and is **not** relied on.
3. **API → Graph:** client-credentials token. The app has `Sites.Selected` with **Write on the AdvisorManagement site only**.
4. **Per-record authorization, as the user** (`backend/src/lib/access.ts`). The API token is exchanged
   **On-Behalf-Of** for a Dataverse token for the same user (`dataverseAsUser` in `graph.ts`), so Dataverse applies
   their security roles, business units, teams and sharing:

   | Operation | The user must (in Dynamics 365, with their own roles) |
   |---|---|
   | List documents | Read the **contact**. Documents whose **source record** they can't read are hidden |
   | Upload | Read the contact; the source record's **Contact Lookup Column** must point to that contact; **write** access to the source record (`CHANGE_ACCESS`) |
   | Preview / download | Read the contact and the document's source record |
   | Delete | Read the contact; **write** access to the source record (or to the contact for contact-level documents) |
   | Any file without `ContactGuid` | Refused: its access can't be verified |

   - "Read" = `GET /<entityset>(id)` as the user succeeds (403/404 → denied). "Write" = `RetrievePrincipalAccess`
     includes `WriteAccess`. The record's lookup column comes from its Document Manager Configuration row.
   - The list filter checks source records in batches (`Microsoft.Dynamics.CRM.In`, 50 ids per call). A document
     whose source record was deleted is hidden too: Dataverse can't tell "deleted" from "not visible to you".
   - Results are cached per request only, so a role change applies on the user's next action. Table metadata
     (entity set name, primary key) is cached per worker.
   - Can't be turned off in production: `ACCESS_CHECKS=off` is honoured only together with `REQUIRE_AUTH=false`
     (local development).
5. **Delete guards:** files only (folders → 400), plus the `X-Confirm-Delete` header must equal the current file
   name (→ 409 otherwise). Deleted files go to the SharePoint recycle bin.

## 7. Invariants (do not break)

- One folder per contact, keyed by GUID suffix. No per-entity folders.
- The client never sends a folder path. The server derives it from `contactId` + `category`.
- Every file created by the control has all four metadata columns, or the file is removed.
- No secrets in the control, its manifest properties, or the repo.
- Every document operation goes through `AccessChecker` (`accessFor(caller)`) before touching SharePoint. Never add
  an endpoint that reads or changes files without it.
