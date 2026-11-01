import { ICategory } from "../categories";

/**
 * Reads the control's configuration from Dataverse as the signed-in user:
 *  - environment variables ce_DocManager* (API URL, tenant, SPA client id, API scope)
 *  - ce_documentmanagerconfig for the hosting table (contact lookup column, allowed categories, size limit)
 *  - ce_documentcategory (labels/colours for every category, including inactive ones for old documents)
 * Results are cached per page load so several controls/forms don't repeat the queries.
 */
export interface IApiSettings {
    apiBaseUrl: string;
    tenantId: string;
    clientId: string;
    apiScope: string;
}

export interface ITableSettings {
    tableLogicalName: string;
    name: string;
    contactLookupColumn: string;
    maxUploadMb?: number;
    /** Active categories allowed on this table, in display order. */
    categories: ICategory[];
}

export const ENV_VARS: Record<keyof IApiSettings, string> = {
    apiBaseUrl: "ce_DocManagerApiBaseUrl",
    tenantId: "ce_DocManagerTenantId",
    clientId: "ce_DocManagerClientId",
    apiScope: "ce_DocManagerApiScope",
};

export class SettingsError extends Error {}

interface DvCategory { ce_name?: string; ce_value?: string; ce_color?: string | null; ce_sortorder?: number | null; statecode?: number }

const toCategory = (c: DvCategory): ICategory => ({
    value: c.ce_value ?? "", label: c.ce_name || c.ce_value || "", color: c.ce_color || undefined, sortOrder: c.ce_sortorder ?? 1000,
});
const bySort = (a: ICategory, b: ICategory) => a.sortOrder - b.sortOrder || a.label.localeCompare(b.label);

const MISSING_TABLES_HINT = "Make sure the Document Manager solution is installed (dataverse/provision.mjs) "
    + "and that your security role includes \"Document Manager Reader\".";

const cache = new Map<string, Promise<unknown>>();
function once<T>(key: string, load: () => Promise<T>): Promise<T> {
    let p = cache.get(key) as Promise<T> | undefined;
    if (!p) {
        p = load().catch(e => { cache.delete(key); throw e; }); // retry on next mount after a failure
        cache.set(key, p);
    }
    return p;
}

export class SettingsService {
    constructor(private webAPI: ComponentFramework.WebApi) {}

    /** Environment variable values (current value, else default), with optional per-form overrides on top. */
    loadApiSettings(overrides: Partial<IApiSettings>): Promise<IApiSettings> {
        return once("env", async () => {
            const names = Object.values(ENV_VARS);
            const r = await this.webAPI.retrieveMultipleRecords("environmentvariabledefinition",
                `?$select=schemaname,defaultvalue&$filter=${names.map(n => `schemaname eq '${n}'`).join(" or ")}`
                + `&$expand=environmentvariabledefinition_environmentvariablevalue($select=value)`)
                .catch(e => { throw new SettingsError(`Could not read environment variables: ${(e as Error).message}`); });
            const values: Record<string, string> = {};
            for (const d of r.entities as { schemaname: string; defaultvalue?: string; environmentvariabledefinition_environmentvariablevalue?: { value?: string }[] }[]) {
                values[d.schemaname.toLowerCase()] = (d.environmentvariabledefinition_environmentvariablevalue?.[0]?.value ?? d.defaultvalue ?? "").trim();
            }
            return values;
        }).then(values => {
            const settings = {} as IApiSettings;
            const missing: string[] = [];
            (Object.keys(ENV_VARS) as (keyof IApiSettings)[]).forEach(k => {
                settings[k] = (overrides[k] || values[ENV_VARS[k].toLowerCase()] || "").trim();
                if (!settings[k]) missing.push(ENV_VARS[k]);
            });
            if (missing.length) {
                throw new SettingsError(`The Document Manager isn't connected yet. Set these environment variables: ${missing.join(", ")}.`);
            }
            return settings;
        });
    }

    /** Configuration row for the hosting table, or null if the table isn't configured. */
    loadTable(tableLogicalName: string): Promise<ITableSettings | null> {
        const t = tableLogicalName.toLowerCase();
        if (!/^[a-z][a-z0-9_]{0,127}$/.test(t)) return Promise.resolve(null);
        return once(`table:${t}`, async () => {
            const r = await this.webAPI.retrieveMultipleRecords("ce_documentmanagerconfig",
                `?$select=ce_name,ce_tablelogicalname,ce_contactlookupcolumn,ce_maxuploadmb`
                + `&$filter=ce_tablelogicalname eq '${t}' and statecode eq 0`
                + `&$expand=ce_documentmanagerconfig_documentcategory($select=ce_name,ce_value,ce_color,ce_sortorder,statecode)`)
                .catch(e => { throw new SettingsError(`Could not read Document Manager configuration: ${(e as Error).message}. ${MISSING_TABLES_HINT}`); });
            const row = r.entities[0] as {
                ce_name?: string; ce_tablelogicalname?: string; ce_contactlookupcolumn?: string; ce_maxuploadmb?: number | null;
                ce_documentmanagerconfig_documentcategory?: DvCategory[];
            } | undefined;
            if (!row) return null;
            return {
                tableLogicalName: t,
                name: row.ce_name || t,
                contactLookupColumn: (row.ce_contactlookupcolumn ?? "").toLowerCase(),
                maxUploadMb: row.ce_maxuploadmb ?? undefined,
                categories: (row.ce_documentmanagerconfig_documentcategory ?? [])
                    .filter(c => (c.statecode ?? 0) === 0 && c.ce_value)
                    .map(toCategory)
                    .sort(bySort),
            };
        });
    }

    /** Every category (active and inactive), for labels and colours of existing documents. */
    loadCatalog(): Promise<ICategory[]> {
        return once("catalog", async () => {
            const r = await this.webAPI.retrieveMultipleRecords("ce_documentcategory",
                "?$select=ce_name,ce_value,ce_color,ce_sortorder,statecode")
                .catch(e => { throw new SettingsError(`Could not read document categories: ${(e as Error).message}. ${MISSING_TABLES_HINT}`); });
            return (r.entities as DvCategory[]).filter(c => c.ce_value).map(toCategory).sort(bySort);
        });
    }
}
