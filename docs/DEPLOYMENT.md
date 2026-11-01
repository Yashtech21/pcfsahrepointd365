# Deployment Guide

End-to-end setup for the **Categorized Document Manager**. Do the steps in order.

```
D365 form ──► PCF control ──(user token, MSAL SSO)──► Azure Function (Easy Auth) ──(app-only, Sites.Selected)──► Microsoft Graph ──► SharePoint /sites/AdvisorManagement/Contacts
```

## 1. Entra ID app registrations

You need **two** registrations.

### 1a. `D365-SP-DocManager` (backend / Graph)
1. Entra ID → App registrations → New registration → name `D365-SP-DocManager`.
2. **API permissions** → Microsoft Graph → *Application* → `Sites.Selected` → **Grant admin consent**.
3. **Certificates & secrets** → new client secret (store it in Key Vault).
4. **Expose an API** → set Application ID URI (`api://<client-id>`) → add scope `user_impersonation`
   (admins and users can consent).
5. *(Optional)* **App roles** → create `DocManager.User` (allowed member types: Users/Groups), assign it to the
   D365 users who should use the control, then set `REQUIRED_ROLE=DocManager.User` on the Function App.

### 1b. `D365-SP-DocManager-Client` (SPA used by the PCF control)
1. New registration → platform **Single-page application** → redirect URI = your org URL,
   e.g. `https://contoso.crm.dynamics.com` (add one per environment).
2. **API permissions** → *My APIs* → `D365-SP-DocManager` → `user_impersonation` → Grant admin consent.

## 2. SharePoint

```powershell
./sharepoint/Setup-ContactsLibrary.ps1 `
    -SiteUrl https://contoso.sharepoint.com/sites/AdvisorManagement `
    -PnPClientId <your PnP login app id> `
    -BackendAppId <D365-SP-DocManager client id>
```

This creates the `Contacts` library (if missing), the columns `DocCategory` (choice), `SourceEntity`,
`SourceRecordId`, `ContactGuid`, indexes `DocCategory`/`SourceRecordId`/`ContactGuid`, and grants the backend
app **Write** on that one site only (Sites.Selected).

## 3. Azure Function (`/backend`)

```bash
cd backend
npm ci
npm run build
func azure functionapp publish <function-app-name>
```

Function App (Node 20, Functions v4) **App settings**:

| Setting | Value |
|---|---|
| `GRAPH_TENANT_ID` | tenant GUID |
| `GRAPH_CLIENT_ID` | `D365-SP-DocManager` client id |
| `GRAPH_CLIENT_SECRET` | `@Microsoft.KeyVault(SecretUri=...)` |
| `SP_HOSTNAME` | `contoso.sharepoint.com` |
| `SP_SITE_PATH` | `/sites/AdvisorManagement` |
| `SP_LIBRARY_NAME` | `Contacts` |
| `REQUIRE_EASY_AUTH` | `true` |
| `REQUIRED_ROLE` | *(optional)* `DocManager.User` |
| `MAX_UPLOAD_MB` | *(optional, default 100)* |

**Authentication** (Function App → Authentication):
- Add identity provider **Microsoft**, pick the existing `D365-SP-DocManager` registration.
- *Allowed token audiences*: `api://<D365-SP-DocManager client id>`.
- *Unauthenticated requests*: **HTTP 401**.

**CORS** (Function App → CORS): add every D365 org URL, e.g. `https://contoso.crm.dynamics.com`.

## 4. PCF control (`/pcf`)

```bash
cd pcf
npm ci
npm run build:prod
```

Package as a solution with the Power Platform CLI:

```bash
mkdir solution && cd solution
pac solution init --publisher-name ContosoEnterprise --publisher-prefix ce
pac solution add-reference --path ../pcf
dotnet build -c Release          # or: msbuild /t:build /restore /p:Configuration=Release
```

Import `solution/bin/Release/solution.zip` into Dataverse. For quick dev iterations: `pac pcf push --publisher-prefix ce`.

## 5. Form configuration

On the **Opportunity** and **Advisor Onboarding** main forms:
1. Add a tab/section called **Upload & Document Manager**.
2. Add the Contact lookup column (`parentcontactid` on Opportunity, your recruit lookup on Advisor Onboarding).
3. Components → **+ Component** → *Categorized Document Manager*, and set:
   - **Document API Base URL**: `https://<function-app>.azurewebsites.net/api`
   - **Entra Tenant ID**: tenant GUID
   - **SPA Client ID**: `D365-SP-DocManager-Client` client id
   - **API Scope**: `api://<D365-SP-DocManager client id>/user_impersonation`
   - **Category Config** *(optional)*: JSON override, see README
4. Hide the label, set the column to span the full width, save and **publish**.

## Troubleshooting

| Symptom | Cause / fix |
|---|---|
| Popup asking to sign in each session | Third-party cookies blocked or SPA redirect URI missing for this org URL. |
| `401` from the API | Easy Auth audience doesn't match the scope configured in the control. |
| `403 Caller lacks the required role` | User isn't assigned the `DocManager.User` app role. |
| `SharePoint error: Access denied` | Sites.Selected grant missing (re-run step 2). |
| Documents uploaded directly in SharePoint don't appear | They lack `ContactGuid`; fill the column in SharePoint. |
