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

### Category → folder (server is the source of truth: `backend/src/lib/categories.ts`)

| `DocCategory` value | Label in UI | Folder | Offered on |
|---|---|---|---|
| `Offer Letter` | Offer Letter | `Offer Letters/` | opportunity |
| `Commissions` | Commissions | `Commissions/` | opportunity |
| `Payslip` | Payslip | `Onboarding Documents/Payslips/` | any other entity |
| `Resignation Letter` | Resignation Letter | `Onboarding Documents/Resignation Letters/` | any other entity |
| `FinCEN_U4` | FinCEN / U4 Verification | `Onboarding Documents/FinCEN_U4/` | any other entity |
| `Misc` | Miscellaneous | `Miscellaneous/` | all |

"Offered on" is the default and can be overridden per entity with the control's `categoryConfig` JSON property.
Unknown names are ignored on the client and rejected (400) by the server.

## 3. SharePoint columns (library `Contacts`)

| Internal name | Type | Indexed | Value written on upload |
|---|---|---|---|
| `DocCategory` | Choice (values above) | yes | category value |
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
4. **Authorization is per user, not per record.** Any authorized caller can list documents for any contact GUID.
   Adding Dataverse record-level checks would go in `functions/documents.ts` (on-behalf-of call to Dataverse).
   This is a known, documented gap; don't assume record-level checks exist.
5. **Delete guards:** files only (folders → 400), plus the `X-Confirm-Delete` header must equal the current file
   name (→ 409 otherwise). Deleted files go to the SharePoint recycle bin.

## 7. Invariants (do not break)

- One folder per contact, keyed by GUID suffix. No per-entity folders.
- The client never sends a folder path. The server derives it from `contactId` + `category`.
- Every file created by the control has all four metadata columns, or the file is removed.
- No secrets in the control, its manifest properties, or the repo.
