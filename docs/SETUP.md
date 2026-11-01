# Setup Guide

End-to-end setup for the **Categorized Document Manager**. First complete the checklist in
[PREREQUISITES.md](PREREQUISITES.md), then do these steps **in order**. Each step ends with a check you can run
before moving on.

```
D365 form ──► PCF control ──(user token, MSAL SSO)──► Azure Function ──(app-only, Sites.Selected)──► Microsoft Graph ──► SharePoint /sites/AdvisorManagement/Contacts
```

| Step | Where | Who |
|---|---|---|
| 1. Entra ID app registrations | Entra admin center | Application Administrator |
| 2. SharePoint library, columns, grant | SharePoint + PowerShell | SharePoint Administrator |
| 3. Azure Function | Azure portal + CLI | Azure Contributor |
| 4. Build and import the control | Developer machine + Power Apps | System Customizer |
| 5. Dataverse tables, configuration and environment variables | Power Apps + `dataverse/provision.mjs` | System Administrator |
| 6. Add the control to forms | Power Apps form designer | System Customizer |
| 7. Users and security | Entra + D365 admin | Admins |
| 8. Smoke test | D365 | Anyone with access |

---

## 1. Entra ID app registrations

You need **two** registrations.

### 1a. `D365-SP-DocManager` (backend, calls Graph)

1. **Entra admin center → App registrations → New registration**. Name: `D365-SP-DocManager`. Single tenant.
   Note the **Application (client) ID** and **Directory (tenant) ID**.
2. **API permissions → Add a permission → Microsoft Graph → Application permissions → `Sites.Selected`**, then
   **Grant admin consent**. Don't add `Sites.ReadWrite.All` or `Files.ReadWrite.All`.
3. **Certificates & secrets → New client secret.** Copy the value straight into Key Vault (step 3). Note the expiry
   date and set a reminder to rotate it.
4. **Expose an API → Set** Application ID URI to `api://<backend-client-id>` → **Add a scope**
   `user_impersonation` (Admins and users; display name "Access document API").
5. *(Optional, recommended)* **App roles → Create app role** `DocManager.User` (Allowed member types: Users/Groups).
   You'll assign it in step 7 and set `REQUIRED_ROLE` in step 3.

### 1b. `D365-SP-DocManager-Client` (used by the control in the browser)

1. **New registration** named `D365-SP-DocManager-Client`, single tenant.
2. **Authentication → Add a platform → Single-page application.** Redirect URI = your org URL, e.g.
   `https://contoso.crm.dynamics.com`. Add **one per environment** (dev, test, prod).
3. **API permissions → Add a permission → My APIs → `D365-SP-DocManager` → Delegated → `user_impersonation`**, then
   **Grant admin consent**.

✅ **Check:** the backend app shows `Sites.Selected` (Application, Granted). The client app shows
`user_impersonation` (Delegated, Granted) and has an SPA redirect URI for each org.

---

## 2. SharePoint

### 2a. Library, columns and grant (scripted, recommended)

```powershell
./sharepoint/Setup-ContactsLibrary.ps1 `
    -SiteUrl https://contoso.sharepoint.com/sites/AdvisorManagement `
    -PnPClientId <your PnP login app id> `
    -BackendAppId <D365-SP-DocManager client id>
```

The script creates the `Contacts` library if it's missing. It then adds the four columns with the correct
**internal names**, makes sure they're **not required**, indexes `DocCategory`, `SourceRecordId` and
`ContactGuid`, and grants the backend app **Write** on this site only.

### 2b. Manual alternative

If you can't run the script, follow PREREQUISITES §3 exactly:
1. Create each column with the **internal** name (`DocCategory`, `SourceEntity`, `SourceRecordId`, `ContactGuid`),
   then rename the display name. All four are **Single line of text** (categories live in Dataverse, so
   `DocCategory` isn't a Choice).
2. **Library settings → Indexed columns → Create a new index** for `DocCategory`, `SourceRecordId`, `ContactGuid`.
3. **Library settings → Versioning settings**: *Require content approval* = **No**, *Require check out* = **No**.
4. Grant Sites.Selected write. A SharePoint admin runs, e.g.:
   `Grant-PnPAzureADAppSitePermission -AppId <backend-client-id> -DisplayName D365-SP-DocManager -Site <site-url> -Permissions Write`

### 2c. Verify

```powershell
./sharepoint/Test-ContactsLibrary.ps1 `
    -SiteUrl https://contoso.sharepoint.com/sites/AdvisorManagement `
    -PnPClientId <your PnP login app id> `
    -BackendAppId <D365-SP-DocManager client id>
```

✅ **Check:** every line says `PASS`. Fix any `FAIL` before continuing.

---

## 3. Azure Function (`/backend`)

1. **Create** (portal or CLI): a Function App with runtime **Node.js 20**, Functions **v4**, *HTTPS only* on. Also
   create a Storage account, a **Key Vault** and **Application Insights**.
2. **Key Vault:** add secret `GraphClientSecret` = the value from step 1a.3. Enable the Function App's
   **system-assigned managed identity** and give it **Get** on secrets (role *Key Vault Secrets User*).
3. **Configuration → Application settings:**

| Setting | Value |
|---|---|
| `GRAPH_TENANT_ID` | tenant GUID |
| `GRAPH_CLIENT_ID` | `D365-SP-DocManager` client id |
| `GRAPH_CLIENT_SECRET` | `@Microsoft.KeyVault(SecretUri=https://<vault>.vault.azure.net/secrets/GraphClientSecret/)` |
| `SP_HOSTNAME` | `contoso.sharepoint.com` |
| `SP_SITE_PATH` | `/sites/AdvisorManagement` |
| `SP_LIBRARY_NAME` | `Contacts` |
| `DATAVERSE_URL` | `https://contoso.crm.dynamics.com` (configuration source; the app user is created in 5c) |
| `CONFIG_CACHE_SECONDS` | *(optional, default 300)* how long table configuration is cached |
| `API_AUDIENCE` | `api://<backend-client-id>,<backend-client-id>` (accepted token audiences) |
| `REQUIRE_AUTH` | `true` (set `false` only for local development) |
| `REQUIRED_ROLE` | *(optional)* `DocManager.User` |
| `MAX_UPLOAD_MB` | *(optional, default 100)* |
| `BLOCKED_EXTENSIONS` | *(optional)* comma-separated list that **replaces** the default blocked types in `fileRules.ts` |
| `GRAPH_TIMEOUT_MS` | *(optional, default 30000)* per-request Graph timeout |

4. **CORS:** add every D365 org URL (e.g. `https://contoso.crm.dynamics.com`). Don't use `*`. The control sends
   `Authorization` and `X-Confirm-Delete`, so a proxy or APIM in front must allow both headers.
5. **Authentication (recommended second layer):** add identity provider **Microsoft** using the existing
   `D365-SP-DocManager` registration. *Allowed token audiences*: `api://<backend-client-id>`. *Unauthenticated
   requests*: **HTTP 401**. The Function verifies tokens itself either way.
6. **Deploy:**

```bash
cd backend
npm ci
npm test
func azure functionapp publish <function-app-name>
```

✅ **Check:** `curl -i https://<function-app>.azurewebsites.net/api/documents?contactId=00000000-0000-0000-0000-000000000000`
returns **401** (no token). Not 404 (wrong URL) and not 500 (bad settings: check Application Insights).

---

## 4. Build and import the control

```bash
cd pcf
npm ci
npm run build:prod
```

Package it as a solution:

```bash
mkdir ../solution && cd ../solution
pac solution init --publisher-name ContosoEnterprise --publisher-prefix ce
pac solution add-reference --path ../pcf
dotnet build -c Release          # or: msbuild /t:build /restore /p:Configuration=Release
```

Import it into Dataverse:
1. **make.powerapps.com → select the environment → Solutions → Import solution** →
   `solution/bin/Release/solution.zip` → **Import**.
2. Use the **unmanaged** solution in dev and a **managed** one in test and production
   (`dotnet build -c Release /p:SolutionPackageType=Managed`).

For quick dev iterations you can instead run `pac auth create --url https://<org>.crm.dynamics.com` once, then
`pac pcf push --publisher-prefix ce` from `pcf/`.

✅ **Check:** **Solutions → your solution → Objects → Custom controls** lists
`EnterpriseDocManagement.CategorizedDocManager`.

---

## 5. Dataverse tables, configuration and environment variables

### 5a. Hosting tables

**Opportunity:** nothing to create. It has the OOB **Contact** lookup (`parentcontactid`).

**Advisor Onboarding** (skip if it already exists with a Contact lookup):
1. **Solutions → your solution → New → Table → Table** → Display name *Advisor Onboarding*
   (logical name becomes e.g. `ce_advisoronboarding`). Save.
2. **Columns → New column** → Display name *Recruit* → Data type **Lookup** → Related table **Contact** → Save
   (logical name e.g. `ce_recruitid`).
3. Add the table to your model-driven app (**App designer → Pages → Add page → Dataverse table**).

**Any other table** where you want the control needs a **Lookup → Contact** column the same way.

✅ **Check:** you can create an Advisor Onboarding record and pick a Contact in *Recruit*.

### 5b. Configuration tables, environment variables, role and seed data

Run as **your own** System Administrator account (details in [CONFIGURATION.md](CONFIGURATION.md)). First check
`dataverse/seed.json`: `tableLogicalName` and `contactLookupColumn` must match your real tables, e.g.
`ce_advisoronboarding` / `ce_recruitid`.

```bash
az login
node dataverse/provision.mjs all --org https://contoso.crm.dynamics.com --dry-run   # review
node dataverse/provision.mjs all --org https://contoso.crm.dynamics.com             # apply
```

This creates the **Document Category** and **Document Manager Configuration** tables, the environment variables
`ce_DocManagerApiBaseUrl`, `ce_DocManagerTenantId`, `ce_DocManagerClientId` and `ce_DocManagerApiScope` (no values
yet), the **Document Manager Reader** role, and the starting categories and table rows.

### 5c. Connect it (once the app registrations and Function App exist)

```bash
node dataverse/provision.mjs set-env --org https://contoso.crm.dynamics.com \
  --api-base-url https://<function-app>.azurewebsites.net/api --tenant-id <tenant-guid> \
  --client-id <spa-client-id> --api-scope api://<backend-client-id>/user_impersonation
node dataverse/provision.mjs app-user --org https://contoso.crm.dynamics.com --app-id <backend-client-id>
```

Then set `DATAVERSE_URL` on the Function App (step 3), so the backend reads the same configuration.

✅ **Check:** in the app, **Document Manager Configurations** has a row per table with the right **Contact Lookup
Column**, each with its Allowed Categories. **Solutions → DocumentManager → Environment variables** shows four
variables with current values.

---

## 6. Add the control to the forms

Repeat for the **Opportunity** main form and the **Advisor Onboarding** main form.

1. **Solutions → your solution → Tables → (table) → Forms → (Main form) → Edit** (modern form designer).
2. Keep (or add) the normal Contact lookup somewhere users can edit it, e.g. the *Summary* tab. The control
   replaces the lookup UI where it's placed, so users need this second copy to choose the contact.
3. **+ Add → Tab → One column**. Name it **Upload & Document Manager**.
4. From **Table columns**, drag the Contact lookup (`parentcontactid` on Opportunity, *Recruit* on Advisor
   Onboarding) into the new tab's section. Yes, the same column a second time.
5. Select that new field. In the right pane: **Components → + Component**. If *Categorized Document Manager* isn't
   listed, choose **Get more components**, find it on the *Build by my org* tab, **Add** it, then select it.
6. Properties: **Contact Lookup** is already bound to the column you dropped. **Leave the four connection
   properties empty**; they come from the environment variables. Fill them only to override one form, e.g. while
   testing against a dev API.

7. Tick **Web**, **Mobile** and **Tablet** (or just Web). Click **Done**.
8. With the field selected: **Properties → Hide label** = on. Set the section/field to use the full width.
9. *(Recommended)* Remove or hide the OOB **Documents** related tab so users have one place for files.
10. **Save and publish.**

✅ **Check:** open an existing Opportunity that has a Contact. The new tab shows skeleton cards, then either
documents or "No documents found". You should **not** see:
- "isn't connected yet": environment variables are missing (5c);
- "Not set up for this table": no active configuration row for the table (5b);
- a lookup mismatch warning: the form's bound column differs from the row's Contact Lookup Column;
- a sign-in popup every time: the SPA redirect URI is missing for this org (1b).

---

## 7. Users and security

1. **D365:** users need Read on the hosting table and on Contact. Standard roles already include Read on their own
   User record. See PREREQUISITES §4.4.
2. **Entra (if `REQUIRED_ROLE` is set):** **Enterprise applications → D365-SP-DocManager → Users and groups → Add**
   and assign `DocManager.User` to the D365 user group.
3. **SharePoint (only for "Open in SharePoint"):** give users Read on the AdvisorManagement site, or leave it off if
   they should only work through D365.

---

## 8. Smoke test

Run this on a test Opportunity whose Contact is "Test Advisor":

| # | Action | Expected |
|---|---|---|
| 1 | Open the record → *Upload & Document Manager* tab | Skeletons, then the grid. Tab counts appear |
| 2 | Pick **Offer Letter**, drop a small PDF, click **Upload** | Progress bar → "1 file(s) uploaded as "Offer Letter"" → card with **New** badge. Counts go up by 1 |
| 3 | SharePoint → `Contacts/Test Advisor_<guid>/Offer Letters/` | The file is there with all four columns filled in |
| 4 | Upload a file over 4 MB | Progress bar moves smoothly. The file arrives intact (it went through the chunked upload) |
| 5 | Try uploading `test.exe`, or a `.txt` renamed to `.pdf` | Rejected with a clear message. Nothing is created in SharePoint |
| 6 | **Preview** a PDF and a DOCX | The preview pane renders both |
| 7 | **Download** | The file downloads |
| 8 | **Delete** → the button is disabled until the box is ticked → confirm | The card disappears. The file is in the SharePoint recycle bin |
| 9 | Open an Advisor Onboarding record for the same contact → *All Contact Documents* | The Opportunity file is listed with Source "Opportunity" |
| 10 | Create a new Opportunity (unsaved) | "Save this record before uploading documents." Upload is disabled |

---

## Troubleshooting

See [TROUBLESHOOTING.md](TROUBLESHOOTING.md) for the full list. The most common setup problems:

| Symptom | Cause / fix |
|---|---|
| "The Document Manager isn't connected yet. Set these environment variables: …" | Run `provision.mjs set-env` (5c) or set the values in the solution |
| "Not set up for this table" | Add or activate a Document Manager Configuration row for that table (5b) |
| "Could not read Document Manager configuration" | Solution not provisioned (5b), or the user lacks the **Document Manager Reader** role |
| Upload: `400 Category "x" is not allowed on …` | Add the category to that table's Allowed Categories |
| Upload: `Dataverse error: …` / `502` | Backend can't read configuration: check `DATAVERSE_URL` and the application user (5c) |
| Sign-in popup every session | SPA redirect URI missing for this org URL (1b.2) or third-party cookies blocked |
| `401 Invalid or expired token` | `API_AUDIENCE` doesn't match the control's *API Scope* |
| `403 Caller lacks the required role` | User isn't assigned `DocManager.User` (step 7.2) |
| `SharePoint error: Access denied` | Sites.Selected grant missing. Re-run step 2 |
| Upload works but files never appear in the list | Column internal names are wrong or `ContactGuid` isn't indexed. Run `Test-ContactsLibrary.ps1` |
| Files stuck "Checked out" in SharePoint | *Require check out* is on, or a column is marked Required (PREREQUISITES §3.3) |
| CORS error in the browser console | Org URL missing from Function App CORS (step 3.4) |
