# Dynamics 365 – SharePoint Categorized Document Manager (PCF)

A Power Apps Component Framework control for Dynamics 365 model-driven apps that replaces the native
SharePoint document grid with:

- **One folder per Contact**: `/Contacts/{FullName}_{ContactGUID}/{Category folder}/`, used by every related record
  (Opportunity, Advisor Onboarding, …), instead of one folder per record.
- **Category picked at upload**, written to SharePoint metadata (`DocCategory`, `SourceEntity`, `SourceRecordId`, `ContactGuid`).
- **Two views**: *This Record's Documents* and *All Contact Documents* (across all entities).
- **Inline preview** (Graph `driveItem: preview` embed), download and delete.
- Drag-and-drop multi-file upload. Files over 4 MB use Graph upload sessions in 5 MB chunks.

![All contact documents with preview](docs/screenshot-all-contact-documents.png)

## Repository layout

| Path | What |
|---|---|
| `pcf/` | The PCF control (React 16.14 + Fluent UI v9 platform libraries, TypeScript) |
| `backend/` | Azure Function (Node 20, TypeScript) that talks to SharePoint via Microsoft Graph with an app-only, `Sites.Selected` token |
| `sharepoint/Setup-ContactsLibrary.ps1` | PnP PowerShell: library columns, indexes, Sites.Selected grant |
| `docs/DEPLOYMENT.md` | Step-by-step deployment (Entra, SharePoint, Function, solution, forms) |

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
Category → folder mapping is enforced server-side in `backend/src/lib/categories.ts`, so a new category must be
added there (and as a choice in SharePoint).

## Security model

- The browser never holds SharePoint/Graph credentials. The control signs the current D365 user in silently (MSAL `ssoSilent`)
  and calls the Function with a user token for the backend's own API scope.
- The Function is protected by App Service Authentication. In code it also checks the tenant and, if configured, a required app role.
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
