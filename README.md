# Dynamics 365 – SharePoint Categorized Document Manager (PCF)

A Power Apps Component Framework control for Dynamics 365 model-driven apps that replaces the native
SharePoint document grid with:

- **One folder per Contact**: `/Contacts/{FullName}_{ContactGUID}/{Category folder}/`, used by every related record
  (Opportunity, Advisor Onboarding, …), instead of one folder per record.
- **Category picked at upload**, written to SharePoint metadata (`DocCategory`, `SourceEntity`, `SourceRecordId`, `ContactGuid`).
- **Two views**: *This Record's Documents* and *All Contact Documents* (across all entities).
- **Inline preview** (Graph `driveItem: preview` embed) and download.
- **Delete with explicit confirmation**: a dialog shows the file, category and source, and the user must tick
  "I understand…" before Delete is enabled. The server also requires the confirmed file name
  (`X-Confirm-Delete`) and refuses to delete folders.
- **Resilient loading**: card skeletons while loading, timeouts on every call with a Retry banner, and duplicate
  calls are skipped while Dynamics repeatedly calls `updateView`. Upload is locked on the first click.
- Drag-and-drop multi-file upload. Files over 4 MB use Graph upload sessions in 5 MB chunks.

![All contact documents with preview](docs/screenshot-all-contact-documents.png)

## Repository layout

| Path | What |
|---|---|
| `pcf/` | The PCF control (React 16.14 + Fluent UI v9 platform libraries, TypeScript). Production bundle ≈ 330 KB / 80 KB gzipped, of which ≈ 284 KB is MSAL; React/Fluent are supplied by Dynamics and icons are inlined. |
| `backend/` | Azure Function (Node 20, TypeScript) that talks to SharePoint via Microsoft Graph with an app-only, `Sites.Selected` token |
| `sharepoint/Setup-ContactsLibrary.ps1` | PnP PowerShell: library columns, indexes, Sites.Selected grant |
| `docs/DEPLOYMENT.md` | Step-by-step deployment (Entra, SharePoint, Function, solution, forms) |
| `docs/ARCHITECTURE.md`, `docs/API_SPECS.md`, `docs/PCF_STANDARDS.md`, `docs/TROUBLESHOOTING.md` | Design, contracts, coding rules and known issues |
| `AGENTS.md` (+ `CLAUDE.md`) | Instructions for AI coding agents (Claude Code, Cursor, Copilot, Windsurf) |

## Categories

| Hosted on | Dropdown | Folder |
|---|---|---|
| Opportunity | Offer Letter | `Offer Letters/` |
| | Commissions | `Commissions/` |
| | Miscellaneous (`Misc`) | `Miscellaneous/` |
| Any other entity (Advisor Onboarding) | Payslip | `Onboarding Documents/Payslips/` |
| | Resignation Letter | `Onboarding Documents/Resignation Letters/` |
| | FinCEN / U4 Verification (`FinCEN_U4`) | `Onboarding Documents/FinCEN_U4/` |
| | Miscellaneous (`Misc`) | `Miscellaneous/` |

The per-entity lists can be overridden with the control's **Category Config** property, e.g.
`{"opportunity":["Offer Letter","Commissions","Misc"],"ce_advisoronboarding":["Payslip","FinCEN_U4","Misc"]}`.
Category → folder mapping is enforced server-side in `backend/src/lib/categories.ts`, which is the source of truth.
To add a category, add it there, to `ALL_CATEGORIES` in `pcf/CategorizedDocManager/categories.ts` and as a choice in
SharePoint. Unknown names in Category Config are ignored, with a console warning.

## Security model

- The browser never holds SharePoint/Graph credentials. The control signs the current D365 user in silently (MSAL `ssoSilent`)
  and calls the Function with a user token for the backend's own API scope.
- The Function verifies that token in code (signature against the tenant's keys, issuer, audience, expiry, tenant and an
  optional app role). It does not rely on the App Service Authentication header, which can be forged if that setting is off.
- The Function calls Graph app-only with **Sites.Selected**, so it can write only to the AdvisorManagement site.
- ⚠️ Authorization is **per user, not per record**: any user allowed to call the API can list documents for any contact GUID.
  If you need Dataverse record-level security, add a check in `backend/src/functions/documents.ts`
  (e.g. call the Dataverse Web API on behalf of the user to confirm read access to the contact).

## Develop

```bash
cd pcf && npm install && npm run build      # control
cd backend && npm install && npm test       # backend + unit tests
cd backend && cp local.settings.sample.json local.settings.json && npm start   # needs Azure Functions Core Tools
```

See **[docs/DEPLOYMENT.md](docs/DEPLOYMENT.md)** for production setup.
