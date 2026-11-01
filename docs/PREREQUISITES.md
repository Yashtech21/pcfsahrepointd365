# Prerequisites

Everything that must exist **before** the Categorized Document Manager will work. Tick each box, then follow
[SETUP.md](SETUP.md). Values in `code` are the defaults used throughout this repo. If you change one, change it
everywhere it's referenced.

---

## 1. People and permissions needed to do the setup

| Role | Needed for |
|---|---|
| **Entra ID**: Application Administrator (or Global Admin) | Creating the two app registrations and granting admin consent (`Sites.Selected`, API scope) |
| **SharePoint Administrator** | Creating the site/library and granting the backend app write access to the one site (Sites.Selected grant) |
| **Azure**: Contributor on a resource group | Creating the Function App, Storage account, Key Vault and Application Insights |
| **Dynamics 365**: System Administrator | Running `dataverse/provision.mjs` (tables, environment variables, role, application user), importing the solution, editing and publishing forms |

## 2. Tools (on the machine doing the setup)

| Tool | Version | Used for |
|---|---|---|
| Node.js | 20 LTS | Building the control and the backend |
| Power Platform CLI (`pac`) | current | Packaging and pushing the control |
| .NET SDK | 6+ | `dotnet build` of the solution project |
| Azure Functions Core Tools | 4.x | Publishing the backend |
| PowerShell + PnP.PowerShell | PowerShell 7.2+, PnP.PowerShell 2.x | `sharepoint/*.ps1` |
| Azure CLI | current | `az login` so `dataverse/provision.mjs` runs as you (or set `DATAVERSE_TOKEN`) |

---

## 3. SharePoint Online

### 3.1 Site and library

| Item | Required value | Notes |
|---|---|---|
| Site | `https://<tenant>.sharepoint.com/sites/AdvisorManagement` | Any site works; set `SP_HOSTNAME` / `SP_SITE_PATH` to match |
| Document library | **`Contacts`** | Matched by display name or URL name (`SP_LIBRARY_NAME`) |

### 3.2 Columns on the `Contacts` library (exact internal names)

| Internal name | Display name (suggested) | Type | Indexed | Required | Values |
|---|---|---|---|---|---|
| `DocCategory` | Document Category | Single line of text | **Yes** | **No** | A category **Value** from Dataverse (`ce_documentcategory`), e.g. `FinCEN_U4` |
| `SourceEntity` | Source Entity | Single line of text | No | **No** | Written by the backend, e.g. `opportunity`, `ce_advisoronboarding` |
| `SourceRecordId` | Source Record ID | Single line of text | **Yes** | **No** | Lowercase GUID of the D365 record |
| `ContactGuid` | Contact ID | Single line of text | **Yes** | **No** | Lowercase GUID of the D365 Contact |

> ⚠️ **Internal names matter, display names don't.** If you create a column in the SharePoint UI as
> "Document Category", its internal name becomes `Document_x0020_Category` and the backend can't write it.
> Create each column with the internal name first (`DocCategory`), then rename its display name. The setup script
> does this correctly. `sharepoint/Test-ContactsLibrary.ps1` verifies it.

> ℹ️ **`DocCategory` is text, not Choice.** Categories are managed in Dataverse, so adding one never needs a
> SharePoint change. If you created it as a Choice column earlier, change it to *Single line of text*: a Choice
> column rejects categories added later.

> ⚠️ **Indexes must exist before the library passes 5,000 items.** After that, SharePoint may refuse to add them,
> and listing by `ContactGuid` will fail on large libraries.

### 3.3 Library settings that must be OFF

| Setting (Library settings →) | Must be | Why |
|---|---|---|
| Versioning → **Require documents to be checked out** | **No** | Uploaded files would stay checked out to the app and be invisible or locked for others |
| Versioning → **Require content approval** | **No** | Files would sit in "Pending" and might not be served to users |
| Any of the four columns marked **Required** | **No** | The file is created before its metadata is written; required columns leave it checked out |

Versioning (major versions) can stay on.

### 3.4 Folders

No folders need to be created. The backend creates `{Contact Full Name}_{Contact GUID}/` and the category
sub-folders on first upload. If you migrate existing documents in, use that folder pattern **and** fill all four
columns, or the control won't show them.

### 3.5 Access

| Who | Access needed |
|---|---|
| Backend app `D365-SP-DocManager` | **Write** on this site via **Sites.Selected** (not `Sites.ReadWrite.All`) |
| End users | **None for upload, list, download or delete** (all done by the backend). **Read** on the site is needed only for the "Open in SharePoint" button. Confirm during testing whether your tenant's preview links also need it. |

---

## 4. Dynamics 365 / Dataverse

### 4.1 Environment

| Item | Requirement |
|---|---|
| App type | **Model-driven app** (Unified Interface). Canvas apps aren't supported by this build. |
| Code components | Allowed (default for model-driven apps). Import is blocked only if your admins disabled custom components. |
| Solution publisher | One publisher for the control's solution, e.g. prefix **`ce`**. The control namespace is `EnterpriseDocManagement`. |

### 4.2 Tables and columns

| Table | Required column | Notes |
|---|---|---|
| **Contact** (`contact`) | Full Name (OOB) | The folder name comes from the lookup's display value (Full Name) |
| **Opportunity** (`opportunity`) | **Contact** lookup `parentcontactid` (OOB) | Already exists. The control is bound to it |
| **Advisor Onboarding** (custom, e.g. `ce_advisoronboarding`) | A **lookup to Contact**, e.g. `ce_recruitid` ("Recruit") | You must create the table and the lookup if they don't exist |
| Any other table you want the control on | A **lookup to Contact** (any logical name) | Plus a configuration row naming that lookup (below) |

Each table can name its Contact lookup differently. The **Document Manager Configuration** row for the table stores
it in **Contact Lookup Column**.

### 4.2a Configuration (created by `dataverse/provision.mjs`; see [CONFIGURATION.md](CONFIGURATION.md))

| Item | What | Notes |
|---|---|---|
| Table **Document Category** (`ce_documentcategory`) | Name, Value, Folder Path, colour, sort order | Replaces the hard-coded category list |
| Table **Document Manager Configuration** (`ce_documentmanagerconfig`) | One active row per hosting table: Table Logical Name, **Contact Lookup Column**, optional Max Upload, Allowed Categories (N:N) | Without a row the control shows "Not set up for this table" and uploads are refused |
| Environment variables | `ce_DocManagerApiBaseUrl`, `ce_DocManagerTenantId`, `ce_DocManagerClientId`, `ce_DocManagerApiScope` | Values set per environment (`provision.mjs set-env`) |
| Security role **Document Manager Reader** | Org-level Read on both tables and on environment variables | Every control user, and the backend's application user |
| Application user for `D365-SP-DocManager` | Created with `provision.mjs app-user` | Lets the backend read configuration |

### 4.3 Forms

| Requirement | Why |
|---|---|
| The Contact lookup column is on the main form **twice**: once as a normal lookup (so users can pick the contact) and once with the control | The control **replaces** the lookup's UI where it's placed and doesn't let users change the contact |
| The record must be **saved** before uploading | Uploads are tagged with the record's GUID. Unsaved records show "Save this record before uploading documents" |

### 4.4 Security roles (end users)

| Privilege | Level | Why |
|---|---|---|
| Read on the hosting table (Opportunity / Advisor Onboarding) | as per your model | To open the form |
| Read on **Contact** | as per your model | To read the bound lookup |
| Read on **User** (`systemuser`) | User (own record) | The control reads the user's sign-in name for single sign-on. Every standard role has this |
| **Document Manager Reader** role | Organization | Read configuration tables and environment variables |
| *(Optional)* Entra app role `DocManager.User` | — | Only if `REQUIRED_ROLE` is set on the Function App |

> 🔒 **Document access follows each user's Dynamics 365 security.** The backend checks every operation in Dataverse
> as the user (On-Behalf-Of):
>
> | Operation | The user must (in Dynamics 365, with their own roles) |
> |---|---|
> | List documents | Read the **contact**. Documents whose **source record** they can't read are hidden |
> | Upload | Read the contact; the source record's **Contact Lookup Column** must point to that contact; **write** access to the source record (`CHANGE_ACCESS`) |
> | Preview / download | Read the contact and the document's source record |
> | Delete | Read the contact; **write** access to the source record (or to the contact for contact-level documents) |
> | Any file without `ContactGuid` | Refused: its access can't be verified |
>
> So users need their **normal** D365 access to the contact and the hosting records, plus **Write** on a record to
> upload or delete its documents. Set `CHANGE_ACCESS=read` on the Function App to let read-only users upload too.

### 4.5 Recommended (not required)

- Hide or remove the out-of-the-box **Documents** related tab (SharePoint document management) on these forms, so
  users don't have two places for files. The control doesn't use Dataverse Document Locations.

---

## 5. Microsoft Entra ID

| App registration | Type | Must have |
|---|---|---|
| **`D365-SP-DocManager`** (backend) | Web/API | Graph **Application** permission `Sites.Selected` with **admin consent**; **Dynamics CRM → Delegated `user_impersonation`** with **admin consent** (for per-record checks as the user); a client secret (stored in Key Vault); **Expose an API** with App ID URI `api://<client-id>` and scope `user_impersonation`; *(optional)* app role `DocManager.User` |
| **`D365-SP-DocManager-Client`** (control) | **Single-page application** | Redirect URI for **every** D365 org URL that hosts the control (e.g. `https://contoso.crm.dynamics.com`, `https://contoso-test.crm.dynamics.com`); delegated permission to `D365-SP-DocManager` / `user_impersonation` with **admin consent** |

## 6. Azure

| Resource | Requirement |
|---|---|
| Function App | Functions **v4**, **Node 20**, HTTPS only. On the Consumption plan the default timeout (5 min) bounds the largest upload. Use Premium/Flex or raise `functionTimeout` for very large files |
| Storage account | Required by Functions |
| Key Vault | Holds `GRAPH_CLIENT_SECRET`. The Function App's managed identity needs *Get* on secrets |
| CORS | Every D365 org URL that hosts the control |
| Application Insights | Recommended (logs include caller name, item id and deletions) |

## 7. Network (from users' browsers)

Must be reachable: `login.microsoftonline.com` (sign-in), `https://<function-app>.azurewebsites.net` (API), and
`*.sharepoint.com` (preview iframes, downloads). The D365 org's Content Security Policy, if customised, must allow
framing `*.sharepoint.com` / `*.officeapps.live.com`.

---

## 8. Values you'll need to collect

Keep these handy; SETUP.md asks for them.

| Value | Example | Used in |
|---|---|---|
| Tenant ID | `11111111-…` | Function `GRAPH_TENANT_ID`, control *Entra Tenant ID* |
| Backend app client ID | `22222222-…` | Function `GRAPH_CLIENT_ID`, `API_AUDIENCE`, control *API Scope*, SharePoint grant |
| Backend client secret | (Key Vault URI) | Function `GRAPH_CLIENT_SECRET` |
| SPA app client ID | `33333333-…` | Control *SPA Client ID* |
| SharePoint host / site path / library | `contoso.sharepoint.com` / `/sites/AdvisorManagement` / `Contacts` | Function `SP_*` settings, scripts |
| Function base URL | `https://d365-sp-docmanager.azurewebsites.net/api` | Control *Document API Base URL* |
| D365 org URL(s) | `https://contoso.crm.dynamics.com` | SPA redirect URIs, Function CORS |
| Advisor Onboarding table + lookup | `ce_advisoronboarding` / `ce_recruitid` | Form setup, `dataverse/seed.json` (configuration row) |
| Dataverse org URL | `https://contoso.crm.dynamics.com` | `provision.mjs --org`, Function `DATAVERSE_URL` |
