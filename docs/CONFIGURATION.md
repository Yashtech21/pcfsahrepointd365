# Configuration (no-code administration)

Categories, the tables that host the control, each table's **Contact lookup column**, per-table upload limits and
the API connection all live in **Dataverse**. Admins change them in the app. No code change, rebuild or SharePoint
change is needed.

```
Environment variables (per environment)          Dataverse tables (data, editable in the app)
  ce_DocManagerApiBaseUrl                          ce_documentcategory           ce_documentmanagerconfig
  ce_DocManagerTenantId                            ─────────────────────         ────────────────────────────────
  ce_DocManagerClientId              ┌────────────► Name (label)     ◄── N:N ──► Name
  ce_DocManagerApiScope              │             Value (→ SharePoint)          Table Logical Name   (unique)
        │                            │             Folder Path                   Contact Lookup Column
        ▼                            │             Badge Colour, Sort Order      Max Upload (MB)
  PCF control ── reads as the user ──┘             Status (Active/Inactive)      Status (Active/Inactive)
  Backend     ── reads as its application user ──► same tables (server-side enforcement, 5 min cache)
```

## 1. Data model

### Document Category (`ce_documentcategory`, organization-owned)

| Column | Logical name | Required | Meaning |
|---|---|---|---|
| Name | `ce_name` | ✓ | Label shown to users, e.g. "FinCEN / U4 Verification" |
| Value | `ce_value` | ✓ (unique key) | Stored in SharePoint `DocCategory`, e.g. `FinCEN_U4`. **Don't change it once documents use it** |
| Folder Path | `ce_folderpath` | ✓ | Folder under the contact folder, e.g. `Onboarding Documents/FinCEN_U4`. Up to 5 levels. `..` is refused, and each segment is sanitised |
| Badge Colour | `ce_color` | | Hex, e.g. `#ca5010`. Empty → a stable automatic colour. Text colour is chosen automatically for contrast |
| Sort Order | `ce_sortorder` | | Lower first in the dropdown |
| Status | `statecode` | | **Inactive** = can't be chosen for new uploads. Existing documents keep their label |

### Document Manager Configuration (`ce_documentmanagerconfig`, organization-owned)

One **active** row per table that hosts the control.

| Column | Logical name | Required | Meaning |
|---|---|---|---|
| Name | `ce_name` | ✓ | Friendly name, e.g. "Advisor Onboarding" |
| Table Logical Name | `ce_tablelogicalname` | ✓ (unique key) | e.g. `opportunity`, `ce_advisoronboarding` |
| **Contact Lookup Column** | `ce_contactlookupcolumn` | ✓ | Logical name of **that table's** lookup to Contact, e.g. `parentcontactid` on Opportunity, `ce_recruitid` on Advisor Onboarding. The control warns if the form binds a different column. It will also be used for per-record security checks later |
| Max Upload (MB) | `ce_maxuploadmb` | | Optional per-table limit. The backend's `MAX_UPLOAD_MB` still caps it |
| Allowed Categories | N:N `ce_documentmanagerconfig_documentcategory` | | Categories offered on this table (active ones only) |
| Status | `statecode` | | **Inactive** = the control shows "Not set up for this table" and the backend refuses uploads |

### Environment variables (type Text)

| Schema name | Value |
|---|---|
| `ce_DocManagerApiBaseUrl` | `https://<function-app>.azurewebsites.net/api` |
| `ce_DocManagerTenantId` | Entra tenant ID |
| `ce_DocManagerClientId` | Client ID of `D365-SP-DocManager-Client` (SPA) |
| `ce_DocManagerApiScope` | `api://<backend client id>/user_impersonation` |

Each environment (dev, test, prod) has its own values. The control's form properties with the same names are
**optional overrides**: leave them empty to use the environment variables.

### Security role: **Document Manager Reader**

Organization-level **Read** on both configuration tables and on environment variable definitions and values.
Give it to every user of the control, typically by adding it to the teams or roles they already have. Give it to the
backend's application user too. Admins who maintain the configuration need Create/Write on the two tables; System
Customizer and System Administrator already have that.

---

## 2. Provisioning (run once per environment, as yourself)

`dataverse/provision.mjs` creates everything above, using **your own** D365 System Administrator account. It needs
Node 20 and either the Azure CLI (`az login`) or a `DATAVERSE_TOKEN` environment variable. Every command can be
re-run safely: anything that already exists is skipped, and seed data is create-only, so it never overwrites edits
made in the app.

```bash
az login                                   # with your D365 admin account

# 1. Review the plan (nothing is sent):
node dataverse/provision.mjs all --org https://contoso.crm.dynamics.com --offline

# 2. Check against the real org without changing anything:
node dataverse/provision.mjs all --org https://contoso.crm.dynamics.com --dry-run

# 3. Create schema + environment variables + role + seed data:
node dataverse/provision.mjs all --org https://contoso.crm.dynamics.com
```

`all` creates:
1. the publisher `ContosoEnterprise` (prefix `ce`) and the unmanaged solution `DocumentManager`, if missing;
2. both tables, their columns, alternate keys and the N:N relationship, all inside that solution, then publishes;
3. the four environment variable **definitions**, with no values;
4. the **Document Manager Reader** role;
5. the categories and table rows from `dataverse/seed.json`.

**Before running `all`**, edit `dataverse/seed.json` if your Advisor Onboarding table or its Contact lookup has a
different logical name than `ce_advisoronboarding` / `ce_recruitid`. The script validates the file first.

### Later, once the app registrations and Function App exist

```bash
# Fill in the environment variables (values aren't put in the solution; they're per environment):
node dataverse/provision.mjs set-env --org https://contoso.crm.dynamics.com \
  --api-base-url https://d365-sp-docmanager.azurewebsites.net/api \
  --tenant-id <tenant-guid> \
  --client-id <spa-client-id> \
  --api-scope api://<backend-client-id>/user_impersonation

# Create the backend's application user and give it the reader role:
node dataverse/provision.mjs app-user --org https://contoso.crm.dynamics.com --app-id <backend-client-id>
```

You can also set the environment variables in the maker portal: **Solutions → DocumentManager → Environment
variables → (variable) → New value**.

### Moving to test and production

Export the `DocumentManager` solution as **managed** and import it. The environment variable definitions travel with
it, and the import asks for their values. Configuration **rows** (categories and tables) are data, not solution
components. Either re-run `provision.mjs seed` against the target org, or move them with the Configuration
Migration tool.

---

## 3. Everyday administration (no code)

### Add a category
1. **Document Categories → New**: Name, Value (no special characters needed), Folder Path, optional colour and sort order.
2. Open the table configuration row(s) where it should be offered → **Allowed Categories → Add Existing** → pick it.
3. That's it. Users see it the next time they open a form. The backend picks it up within `CONFIG_CACHE_SECONDS`
   (default 5 minutes).

### Retire a category
**Deactivate** it rather than deleting it. It disappears from the dropdowns, and existing documents keep their
label and colour.

### Rename a category
Change **Name** freely. Don't change **Value**: documents already in SharePoint store the old value.

### Put the control on another table
1. Make sure the table has a **lookup to Contact**.
2. **Document Manager Configurations → New**: Table Logical Name, **Contact Lookup Column** (that lookup's logical
   name), optional Max Upload, then **Allowed Categories → Add Existing**.
3. Add the control to the table's main form, bound to that lookup (see SETUP §6). Leave the connection properties
   empty, since they come from the environment variables.

### Turn the control off for a table
Deactivate its configuration row. The control shows "Not set up for this table", and the backend refuses uploads.

---

## 4. Where it's enforced

| Rule | Control (UX) | Backend (enforced) |
|---|---|---|
| Table must have an active configuration row | "Not set up for this table" message | `400 Document management is not configured for table "x"` |
| Category must be active and allowed on that table | Only those are offered | `400 Category "x" is not allowed on <table>` |
| Per-table size limit | Files refused when selected | `413 File is too large (limit N MB for <table>)` |
| Folder path | — | Sanitised. `..`, empty paths or more than 5 levels → `400` |
| Bound lookup = configured Contact Lookup Column | Warning banner | (used by per-record security later) |

The backend caches configuration per table for `CONFIG_CACHE_SECONDS` (default 300). Failed reads aren't
cached. For local development, set `DOC_CONFIG_FILE=../dataverse/seed.json` to read configuration from the seed file
instead of Dataverse.
