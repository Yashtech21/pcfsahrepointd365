#!/usr/bin/env node
/**
 * Provisions the Dataverse side of the Categorized Document Manager, as the signed-in admin.
 *
 *   node dataverse/provision.mjs <command> --org https://<org>.crm.dynamics.com [options]
 *
 * Commands (idempotent; anything that already exists is left as is):
 *   all        schema + env + role + seed (+ publish)                    ← first run
 *   schema     tables, columns, alternate keys, N:N relationship
 *   env        environment variable definitions (no values)
 *   role       "Document Manager Reader" security role (read config tables + env vars)
 *   seed       categories and per-table configuration from dataverse/seed.json (create-only)
 *   set-env    set environment variable values  (--api-base-url --tenant-id --client-id --api-scope)
 *   app-user   create the backend's application user and give it the reader role (--app-id <client id>)
 *
 * Options:
 *   --org <url>            Dataverse org URL (required)
 *   --solution <name>      Unmanaged solution to add components to (default DocumentManager)
 *   --publisher <name>     Publisher unique name, prefix must be "ce" (default ContosoEnterprise)
 *   --seed-file <path>     Default dataverse/seed.json
 *   --dry-run              Do the read-only checks, print the writes, send nothing
 *   --offline              Print the full plan without contacting Dataverse (assumes nothing exists)
 *
 * Authentication: your own account. Either `az login` first (Azure CLI is used to get a token for the org),
 * or set DATAVERSE_TOKEN to a bearer token for the org.
 */
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

// ───────────────────────────── Definitions (single source for the Dataverse schema) ─────────────────────────────

export const PREFIX = "ce";
export const ROLE_NAME = "Document Manager Reader";
export const RELATIONSHIP = "ce_documentmanagerconfig_documentcategory";
const LANG = 1033;

export const label = (text) => ({
    "@odata.type": "Microsoft.Dynamics.CRM.Label",
    LocalizedLabels: [{ "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel", Label: text, LanguageCode: LANG }],
});

const required = (on) => ({ Value: on ? "ApplicationRequired" : "None", CanBeChanged: true, ManagedPropertyLogicalName: "canmodifyrequirementlevelsettings" });

export const stringAttr = (schema, display, maxLength, isRequired, description = "", primary = false) => ({
    "@odata.type": "Microsoft.Dynamics.CRM.StringAttributeMetadata",
    SchemaName: schema,
    DisplayName: label(display),
    Description: label(description),
    RequiredLevel: required(isRequired),
    MaxLength: maxLength,
    FormatName: { Value: "Text" },
    ...(primary ? { IsPrimaryName: true } : {}),
});

export const intAttr = (schema, display, min, max, description = "") => ({
    "@odata.type": "Microsoft.Dynamics.CRM.IntegerAttributeMetadata",
    SchemaName: schema,
    DisplayName: label(display),
    Description: label(description),
    RequiredLevel: required(false),
    Format: "None",
    MinValue: min,
    MaxValue: max,
});

export const TABLES = [
    {
        schema: "ce_DocumentCategory",
        display: "Document Category",
        plural: "Document Categories",
        description: "Categories offered by the Categorized Document Manager. Value is what is stored in SharePoint.",
        primary: stringAttr("ce_Name", "Name", 100, true, "Label shown to users", true),
        attributes: [
            stringAttr("ce_Value", "Value", 100, true, "Stored in the SharePoint DocCategory column. Don't change it once documents use it; deactivate and create a new one instead."),
            stringAttr("ce_FolderPath", "Folder Path", 400, true, "Folder under the contact folder, e.g. Onboarding Documents/Payslips"),
            stringAttr("ce_Color", "Badge Colour", 7, false, "Hex colour for the category badge, e.g. #0f6cbd"),
            intAttr("ce_SortOrder", "Sort Order", 0, 100000, "Lower numbers appear first in the dropdown"),
        ],
        keys: [{ schema: "ce_value_key", display: "Value", attributes: ["ce_value"] }],
    },
    {
        schema: "ce_DocumentManagerConfig",
        display: "Document Manager Configuration",
        plural: "Document Manager Configurations",
        description: "One row per Dataverse table that hosts the Categorized Document Manager.",
        primary: stringAttr("ce_Name", "Name", 100, true, "Friendly name, e.g. Opportunity", true),
        attributes: [
            stringAttr("ce_TableLogicalName", "Table Logical Name", 128, true, "Logical name of the hosting table, e.g. opportunity"),
            stringAttr("ce_ContactLookupColumn", "Contact Lookup Column", 128, true, "Logical name of the lookup to Contact on that table, e.g. parentcontactid"),
            intAttr("ce_MaxUploadMb", "Max Upload (MB)", 1, 2000, "Optional per-table limit; the backend's MAX_UPLOAD_MB still caps it"),
        ],
        keys: [{ schema: "ce_tablelogicalname_key", display: "Table Logical Name", attributes: ["ce_tablelogicalname"] }],
    },
];

export const ENV_VARS = [
    { schema: "ce_DocManagerApiBaseUrl", display: "Document Manager - API Base URL", description: "e.g. https://d365-sp-docmanager.azurewebsites.net/api" },
    { schema: "ce_DocManagerTenantId", display: "Document Manager - Entra Tenant ID", description: "Directory (tenant) ID" },
    { schema: "ce_DocManagerClientId", display: "Document Manager - SPA Client ID", description: "Client ID of the D365-SP-DocManager-Client app registration" },
    { schema: "ce_DocManagerApiScope", display: "Document Manager - API Scope", description: "api://<backend client id>/user_impersonation" },
];
const ENV_FLAGS = { "api-base-url": "ce_DocManagerApiBaseUrl", "tenant-id": "ce_DocManagerTenantId", "client-id": "ce_DocManagerClientId", "api-scope": "ce_DocManagerApiScope" };

export function tableDefinition(t) {
    return {
        "@odata.type": "Microsoft.Dynamics.CRM.EntityMetadata",
        SchemaName: t.schema,
        DisplayName: label(t.display),
        DisplayCollectionName: label(t.plural),
        Description: label(t.description),
        OwnershipType: "OrganizationOwned",
        IsActivity: false,
        HasActivities: false,
        HasNotes: false,
        Attributes: [t.primary],
    };
}

export function relationshipDefinition() {
    const menu = (text) => ({ Behavior: "UseLabel", Group: "Details", Label: label(text), Order: 10000 });
    return {
        "@odata.type": "Microsoft.Dynamics.CRM.ManyToManyRelationshipMetadata",
        SchemaName: RELATIONSHIP,
        Entity1LogicalName: "ce_documentmanagerconfig",
        Entity2LogicalName: "ce_documentcategory",
        IntersectEntityName: RELATIONSHIP,
        Entity1NavigationPropertyName: RELATIONSHIP,
        Entity2NavigationPropertyName: RELATIONSHIP,
        Entity1AssociatedMenuConfiguration: menu("Allowed Categories"),
        Entity2AssociatedMenuConfiguration: menu("Used by Tables"),
    };
}

export const odataString = (s) => `'${String(s).replace(/'/g, "''")}'`;
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** Validates seed.json so a typo fails before anything is written. */
export function validateSeed(seed) {
    const errors = [];
    const values = new Set();
    for (const c of seed.categories ?? []) {
        if (!c.value || !c.name || !c.folderPath) errors.push(`category ${JSON.stringify(c)} needs value, name and folderPath`);
        if (values.has(c.value)) errors.push(`duplicate category value "${c.value}"`);
        values.add(c.value);
        if (c.folderPath && c.folderPath.split("/").some(seg => !seg.trim() || seg === "." || seg === "..")) errors.push(`category "${c.value}" has an invalid folderPath`);
        if (c.color && !/^#[0-9a-f]{6}$/i.test(c.color)) errors.push(`category "${c.value}" colour must look like #0f6cbd`);
    }
    for (const t of seed.tables ?? []) {
        if (!/^[a-z][a-z0-9_]{0,127}$/.test(t.tableLogicalName ?? "")) errors.push(`table "${t.tableLogicalName}" is not a valid logical name`);
        if (!/^[a-z][a-z0-9_]{0,127}$/.test(t.contactLookupColumn ?? "")) errors.push(`table "${t.tableLogicalName}" has an invalid contactLookupColumn`);
        for (const v of t.categories ?? []) if (!values.has(v)) errors.push(`table "${t.tableLogicalName}" references unknown category "${v}"`);
    }
    return errors;
}

// ───────────────────────────── Web API client ─────────────────────────────

export class DataverseError extends Error {
    constructor(status, message) { super(message); this.status = status; }
}

export function createClient({ org, token, fetchImpl = fetch, dryRun = false, offline = false, solution, log = console.log }) {
    const base = `${org.replace(/\/+$/, "")}/api/data/v9.2`;
    let fakeId = 0;

    async function send(method, path, body, { timeoutMs = 120000, inSolution = false, headers = {} } = {}) {
        const url = path.startsWith("http") ? path : `${base}${path}`;
        const write = method !== "GET";
        if (offline || (dryRun && write)) {
            if (write) log(`  [${offline ? "plan" : "dry-run"}] ${method} ${url.replace(base, "")}`);
            if (offline && !write) throw new DataverseError(404, "offline");
            return { id: `00000000-0000-0000-0000-${String(++fakeId).padStart(12, "0")}`, body: {} };
        }
        const res = await fetchImpl(url, {
            method,
            signal: AbortSignal.timeout(timeoutMs),
            headers: {
                Authorization: `Bearer ${token}`,
                Accept: "application/json",
                "OData-Version": "4.0",
                "OData-MaxVersion": "4.0",
                ...(body !== undefined ? { "Content-Type": "application/json; charset=utf-8" } : {}),
                ...(inSolution && solution ? { "MSCRM.SolutionUniqueName": solution } : {}),
                ...headers,
            },
            body: body !== undefined ? JSON.stringify(body) : undefined,
        });
        const text = await res.text();
        let json = {};
        try { json = text ? JSON.parse(text) : {}; } catch { /* non-JSON */ }
        if (!res.ok) throw new DataverseError(res.status, json?.error?.message ?? `${res.status} ${res.statusText}`);
        const entityId = res.headers.get("OData-EntityId") ?? "";
        const id = /\(([0-9a-f-]{36})\)\s*$/i.exec(entityId)?.[1];
        return { id, body: json };
    }

    return {
        base,
        offline,
        get: (p, o) => send("GET", p, undefined, o).then(r => r.body),
        /** Collection query → array. Offline (or a 404 on the collection) yields []. */
        async list(p) {
            try { return (await send("GET", p)).body.value ?? []; } catch (e) { if (e.status === 404) return []; throw e; }
        },
        post: (p, b, o) => send("POST", p, b, o),
        patch: (p, b, o) => send("PATCH", p, b, o),
        async exists(p) {
            try { await send("GET", p); return true; } catch (e) { if (e.status === 404) return false; throw e; }
        },
    };
}

// ───────────────────────────── Steps ─────────────────────────────

async function rootBusinessUnit(c) {
    if (c.offline) return "00000000-0000-0000-0000-0000000000b0";
    const id = (await c.list(`/businessunits?$select=businessunitid&$filter=_parentbusinessunitid_value eq null`))[0]?.businessunitid;
    if (!id) throw new Error("Could not find the root business unit");
    return id;
}

export async function ensureSolution(c, { solution, publisher, log = console.log }) {
    const sol = { value: await c.list(`/solutions?$select=solutionid&$filter=uniquename eq ${odataString(solution)}`) };
    if (sol.value?.length) { log(`✓ Solution ${solution} exists`); return; }
    let pub = { value: await c.list(`/publishers?$select=publisherid,customizationprefix&$filter=uniquename eq ${odataString(publisher)}`) };
    let publisherId = pub.value?.[0]?.publisherid;
    if (publisherId && pub.value[0].customizationprefix !== PREFIX) {
        throw new Error(`Publisher ${publisher} has prefix "${pub.value[0].customizationprefix}"; this solution needs prefix "${PREFIX}"`);
    }
    if (!publisherId) {
        log(`+ Publisher ${publisher} (prefix ${PREFIX})`);
        publisherId = (await c.post(`/publishers`, {
            uniquename: publisher, friendlyname: publisher, customizationprefix: PREFIX, customizationoptionvalueprefix: 72_700,
        })).id;
    }
    log(`+ Solution ${solution}`);
    await c.post(`/solutions`, {
        uniquename: solution, friendlyname: "Categorized Document Manager", version: "1.0.0.0",
        "publisherid@odata.bind": `/publishers(${publisherId})`,
    });
}

export async function ensureSchema(c, { log = console.log } = {}) {
    for (const t of TABLES) {
        const logical = t.schema.toLowerCase();
        if (await c.exists(`/EntityDefinitions(LogicalName='${logical}')?$select=MetadataId`)) {
            log(`✓ Table ${logical} exists`);
        } else {
            log(`+ Table ${logical} (this can take a minute)`);
            await c.post(`/EntityDefinitions`, tableDefinition(t), { inSolution: true, timeoutMs: 600000 });
        }
        for (const a of t.attributes) {
            const al = a.SchemaName.toLowerCase();
            if (await c.exists(`/EntityDefinitions(LogicalName='${logical}')/Attributes(LogicalName='${al}')?$select=MetadataId`)) {
                log(`  ✓ Column ${al}`);
            } else {
                log(`  + Column ${al}`);
                await c.post(`/EntityDefinitions(LogicalName='${logical}')/Attributes`, a, { inSolution: true, timeoutMs: 300000 });
            }
        }
        for (const k of t.keys) {
            if (await c.exists(`/EntityDefinitions(LogicalName='${logical}')/Keys(LogicalName='${k.schema}')?$select=MetadataId`)) {
                log(`  ✓ Key ${k.schema}`);
            } else {
                log(`  + Alternate key ${k.schema} (unique ${k.attributes.join(", ")}; activates in the background)`);
                await c.post(`/EntityDefinitions(LogicalName='${logical}')/Keys`, {
                    SchemaName: k.schema, DisplayName: label(k.display), KeyAttributes: k.attributes,
                }, { inSolution: true });
            }
        }
    }
    if (await c.exists(`/RelationshipDefinitions(SchemaName='${RELATIONSHIP}')?$select=MetadataId`)) {
        log(`✓ Relationship ${RELATIONSHIP}`);
    } else {
        log(`+ Relationship ${RELATIONSHIP} (allowed categories per table)`);
        await c.post(`/RelationshipDefinitions`, relationshipDefinition(), { inSolution: true, timeoutMs: 300000 });
    }
}

export async function ensureEnvVars(c, { log = console.log } = {}) {
    for (const v of ENV_VARS) {
        const r = { value: await c.list(`/environmentvariabledefinitions?$select=environmentvariabledefinitionid&$filter=schemaname eq ${odataString(v.schema)}`) };
        if (r.value?.length) { log(`✓ Environment variable ${v.schema}`); continue; }
        log(`+ Environment variable ${v.schema} (no value yet)`);
        await c.post(`/environmentvariabledefinitions`, {
            schemaname: v.schema, displayname: v.display, description: v.description,
            type: 100000000, // String
        }, { inSolution: true });
    }
}

export async function ensureRole(c, { log = console.log } = {}) {
    const bu = await rootBusinessUnit(c);
    const existing = { value: await c.list(`/roles?$select=roleid&$filter=name eq ${odataString(ROLE_NAME)} and _businessunitid_value eq ${bu}`) };
    let roleId = existing.value?.[0]?.roleid;
    if (roleId) {
        log(`✓ Role "${ROLE_NAME}" exists`);
    } else {
        log(`+ Role "${ROLE_NAME}"`);
        roleId = (await c.post(`/roles`, { name: ROLE_NAME, "businessunitid@odata.bind": `/businessunits(${bu})` }, { inSolution: true })).id;
    }
    const names = ["prvReadce_DocumentCategory", "prvReadce_DocumentManagerConfig", "prvReadEnvironmentVariableDefinition", "prvReadEnvironmentVariableValue"];
    const privileges = [];
    for (const name of names) {
        const p = { value: await c.list(`/privileges?$select=privilegeid,name&$filter=name eq ${odataString(name)}`) };
        const id = p.value?.[0]?.privilegeid;
        if (!id && !c.offline) throw new Error(`Privilege ${name} not found. Run "schema" (and let it publish) before "role".`);
        privileges.push({ Depth: "Global", PrivilegeId: id ?? "00000000-0000-0000-0000-000000000000", BusinessUnitId: bu, PrivilegeName: name });
    }
    log(`  + Read (organization) on ${names.join(", ")}`);
    await c.post(`/roles(${roleId})/Microsoft.Dynamics.CRM.AddPrivilegesRole`, { Privileges: privileges });
    return roleId;
}

async function entitySet(c, logical, fallback) {
    try {
        return (await c.get(`/EntityDefinitions(LogicalName='${logical}')?$select=EntitySetName`)).EntitySetName ?? fallback;
    } catch { return fallback; }
}

export async function seed(c, seedData, { log = console.log } = {}) {
    const errors = validateSeed(seedData);
    if (errors.length) throw new Error(`seed file is invalid:\n  - ${errors.join("\n  - ")}`);
    const catSet = await entitySet(c, "ce_documentcategory", "ce_documentcategories");
    const cfgSet = await entitySet(c, "ce_documentmanagerconfig", "ce_documentmanagerconfigs");
    const catIds = new Map();

    for (const cat of seedData.categories) {
        const r = { value: await c.list(`/${catSet}?$select=ce_documentcategoryid&$filter=ce_value eq ${odataString(cat.value)}`) };
        let id = r.value?.[0]?.ce_documentcategoryid;
        if (id) log(`✓ Category ${cat.value}`);
        else {
            log(`+ Category ${cat.value} → ${cat.folderPath}`);
            id = (await c.post(`/${catSet}`, {
                ce_name: cat.name, ce_value: cat.value, ce_folderpath: cat.folderPath,
                ...(cat.color ? { ce_color: cat.color } : {}), ...(cat.sortOrder !== undefined ? { ce_sortorder: cat.sortOrder } : {}),
            })).id;
        }
        catIds.set(cat.value, id);
    }

    for (const t of seedData.tables) {
        const r = { value: await c.list(`/${cfgSet}?$select=ce_documentmanagerconfigid&$filter=ce_tablelogicalname eq ${odataString(t.tableLogicalName)}&$expand=${RELATIONSHIP}($select=ce_value)`) };
        let id = r.value?.[0]?.ce_documentmanagerconfigid;
        const linked = new Set((r.value?.[0]?.[RELATIONSHIP] ?? []).map(x => x.ce_value));
        if (id) log(`✓ Table configuration ${t.tableLogicalName}`);
        else {
            log(`+ Table configuration ${t.tableLogicalName} (contact lookup: ${t.contactLookupColumn})`);
            id = (await c.post(`/${cfgSet}`, {
                ce_name: t.name ?? t.tableLogicalName, ce_tablelogicalname: t.tableLogicalName, ce_contactlookupcolumn: t.contactLookupColumn,
                ...(t.maxUploadMb ? { ce_maxuploadmb: t.maxUploadMb } : {}),
            })).id;
        }
        for (const v of t.categories) {
            if (linked.has(v)) continue;
            log(`  + Allow ${v} on ${t.tableLogicalName}`);
            await c.post(`/${cfgSet}(${id})/${RELATIONSHIP}/$ref`, { "@odata.id": `${c.base}/${catSet}(${catIds.get(v)})` });
        }
    }
}

export async function setEnvValues(c, values, { log = console.log } = {}) {
    for (const [schema, value] of Object.entries(values)) {
        const def = (await c.list(`/environmentvariabledefinitions?$select=environmentvariabledefinitionid&$filter=schemaname eq ${odataString(schema)}&$expand=environmentvariabledefinition_environmentvariablevalue($select=environmentvariablevalueid)`))[0]
            ?? (c.offline ? { environmentvariabledefinitionid: "00000000-0000-0000-0000-0000000000e0" } : undefined);
        if (!def) throw new Error(`Environment variable ${schema} doesn't exist; run "env" first`);
        const current = def.environmentvariabledefinition_environmentvariablevalue?.[0]?.environmentvariablevalueid;
        if (current) {
            log(`~ ${schema} = ${value}`);
            await c.patch(`/environmentvariablevalues(${current})`, { value });
        } else {
            log(`+ ${schema} = ${value}`);
            // Values are deliberately NOT added to the solution: they differ per environment.
            await c.post(`/environmentvariablevalues`, {
                schemaname: schema, value,
                "EnvironmentVariableDefinitionId@odata.bind": `/environmentvariabledefinitions(${def.environmentvariabledefinitionid})`,
            });
        }
    }
}

export async function ensureAppUser(c, appId, { log = console.log } = {}) {
    if (!GUID.test(appId ?? "")) throw new Error("--app-id must be the backend app registration's client ID (GUID)");
    const bu = await rootBusinessUnit(c);
    let userId = (await c.list(`/systemusers?$select=systemuserid&$filter=applicationid eq ${appId}`))[0]?.systemuserid;
    if (userId) log(`✓ Application user for ${appId} exists`);
    else {
        log(`+ Application user for ${appId}`);
        userId = (await c.post(`/systemusers`, { applicationid: appId, "businessunitid@odata.bind": `/businessunits(${bu})` })).id;
    }
    const roleId = (await c.list(`/roles?$select=roleid&$filter=name eq ${odataString(ROLE_NAME)} and _businessunitid_value eq ${bu}`))[0]?.roleid
        ?? (c.offline ? "00000000-0000-0000-0000-0000000000a0" : undefined);
    if (!roleId) throw new Error(`Role "${ROLE_NAME}" not found; run "role" first`);
    const assigned = await c.list(`/systemusers(${userId})/systemuserroles_association?$select=roleid&$filter=roleid eq ${roleId}`);
    if (assigned.length) { log(`✓ Role "${ROLE_NAME}" already assigned`); return; }
    log(`+ Assign role "${ROLE_NAME}"`);
    await c.post(`/systemusers(${userId})/systemuserroles_association/$ref`, { "@odata.id": `${c.base}/roles(${roleId})` });
}

export async function publish(c, { log = console.log } = {}) {
    log(`+ Publish customizations`);
    const entities = TABLES.map(t => `<entity>${t.schema.toLowerCase()}</entity>`).join("");
    await c.post(`/PublishXml`, { ParameterXml: `<importexportxml><entities>${entities}</entities></importexportxml>` }, { timeoutMs: 600000 });
}

// ───────────────────────────── CLI ─────────────────────────────

export function parseArgs(argv) {
    const [command = "help", ...rest] = argv;
    const opts = { command };
    for (let i = 0; i < rest.length; i++) {
        const a = rest[i];
        if (!a.startsWith("--")) throw new Error(`Unexpected argument ${a}`);
        const key = a.slice(2);
        if (key === "dry-run" || key === "offline") { opts[key] = true; continue; }
        const value = rest[++i];
        if (value === undefined || value.startsWith("--")) throw new Error(`--${key} needs a value`);
        opts[key] = value;
    }
    return opts;
}

function getToken(org) {
    if (process.env.DATAVERSE_TOKEN) return process.env.DATAVERSE_TOKEN;
    try {
        return execFileSync("az", ["account", "get-access-token", "--resource", org, "--query", "accessToken", "-o", "tsv"],
            { encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], shell: process.platform === "win32" }).trim();
    } catch (e) {
        throw new Error(`Could not get a token with Azure CLI. Run "az login" with your D365 admin account, or set DATAVERSE_TOKEN.\n${e.stderr ?? e.message}`);
    }
}

export async function main(argv = process.argv.slice(2), env = {}) {
    const opts = parseArgs(argv);
    const log = env.log ?? console.log;
    if (opts.command === "help" || !opts.org) {
        log(readFileSync(fileURLToPath(import.meta.url), "utf8").split("*/")[0].replace(/^#!.*\n\/\*\*?/, "").replace(/^ \* ?/gm, ""));
        if (opts.command !== "help") throw new Error("--org is required");
        return;
    }
    if (!/^https:\/\/[a-z0-9-]+\.crm[0-9]*\.dynamics\.com\/?$/i.test(opts.org)) throw new Error(`--org must look like https://<org>.crm.dynamics.com (got ${opts.org})`);
    const solution = opts.solution ?? "DocumentManager";
    const publisher = opts.publisher ?? "ContosoEnterprise";
    const offline = !!opts.offline;
    const c = env.client ?? createClient({ org: opts.org, token: offline ? "" : getToken(opts.org), dryRun: !!opts["dry-run"], offline, solution, log });
    const seedFile = opts["seed-file"] ?? resolve(dirname(fileURLToPath(import.meta.url)), "seed.json");

    switch (opts.command) {
        case "all":
            await ensureSolution(c, { solution, publisher, log });
            await ensureSchema(c, { log });
            await publish(c, { log });
            await ensureEnvVars(c, { log });
            await ensureRole(c, { log });
            await seed(c, JSON.parse(readFileSync(seedFile, "utf8")), { log });
            break;
        case "schema": await ensureSolution(c, { solution, publisher, log }); await ensureSchema(c, { log }); await publish(c, { log }); break;
        case "env": await ensureSolution(c, { solution, publisher, log }); await ensureEnvVars(c, { log }); break;
        case "role": await ensureSolution(c, { solution, publisher, log }); await ensureRole(c, { log }); break;
        case "seed": await seed(c, JSON.parse(readFileSync(seedFile, "utf8")), { log }); break;
        case "set-env": {
            const values = {};
            for (const [flag, schema] of Object.entries(ENV_FLAGS)) if (opts[flag]) values[schema] = opts[flag];
            if (!Object.keys(values).length) throw new Error(`set-env needs at least one of ${Object.keys(ENV_FLAGS).map(f => `--${f}`).join(", ")}`);
            await setEnvValues(c, values, { log });
            break;
        }
        case "app-user": await ensureAppUser(c, opts["app-id"], { log }); break;
        default: throw new Error(`Unknown command ${opts.command}`);
    }
    log(offline ? "\nPlan complete (nothing was sent)." : opts["dry-run"] ? "\nDry run complete (no changes were made)." : "\nDone.");
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
    main().catch(e => { console.error(`\n✗ ${e.message}`); process.exit(1); });
}
