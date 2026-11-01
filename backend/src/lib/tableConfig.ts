import { readFileSync } from "node:fs";
import { config } from "./config";
import { dataverse } from "./graph";
import { sanitizeName } from "./naming";

/**
 * Categories and per-table settings come from Dataverse (tables ce_documentcategory and
 * ce_documentmanagerconfig, see dataverse/provision.mjs), so admins add categories or tables
 * without code changes. This module is the backend's single source for them.
 */
export interface CategoryDef {
    value: string;          // stored in SharePoint DocCategory
    name: string;           // label
    folderPath: string;     // under the contact folder
    color?: string;
    sortOrder: number;
}

export interface TableConfig {
    tableLogicalName: string;
    name: string;
    /** Logical name of the Contact lookup on that table (e.g. parentcontactid, ce_recruitid). */
    contactLookupColumn: string;
    maxUploadMb?: number;
    /** Active categories allowed on this table, in display order. */
    categories: CategoryDef[];
}

export interface ConfigProvider {
    getTable(tableLogicalName: string): Promise<TableConfig | undefined>;
}

export class ConfigError extends Error {}

const MAX_FOLDER_DEPTH = 5;

/**
 * Folder paths are admin-entered data: sanitize every segment and refuse anything that could escape
 * the contact folder ("..", empty segments) or nest absurdly deep.
 */
export function safeFolderPath(path: string): string {
    const segments = (path ?? "").split(/[\\/]/).map(s => s.trim()).filter(s => s.length > 0);
    if (!segments.length) throw new ConfigError("Category folder path is empty");
    if (segments.length > MAX_FOLDER_DEPTH) throw new ConfigError(`Category folder path is deeper than ${MAX_FOLDER_DEPTH} levels`);
    if (segments.some(s => s === "." || s === "..")) throw new ConfigError("Category folder path may not contain . or ..");
    return segments.map(sanitizeName).join("/");
}

const bySortThenName = (a: CategoryDef, b: CategoryDef) => a.sortOrder - b.sortOrder || a.name.localeCompare(b.name);

interface DvCategory { ce_name?: string; ce_value?: string; ce_folderpath?: string; ce_color?: string | null; ce_sortorder?: number | null; statecode?: number }
interface DvConfigRow {
    ce_name?: string; ce_tablelogicalname?: string; ce_contactlookupcolumn?: string; ce_maxuploadmb?: number | null;
    ce_documentmanagerconfig_documentcategory?: DvCategory[];
}

/** Maps a Dataverse configuration row (with expanded categories) to TableConfig. Inactive categories are dropped. */
export function fromDataverseRow(row: DvConfigRow): TableConfig {
    const categories = (row.ce_documentmanagerconfig_documentcategory ?? [])
        .filter(c => (c.statecode ?? 0) === 0 && c.ce_value && c.ce_folderpath)
        .map(c => ({
            value: c.ce_value!, name: c.ce_name || c.ce_value!, folderPath: c.ce_folderpath!,
            color: c.ce_color ?? undefined, sortOrder: c.ce_sortorder ?? 1000,
        }))
        .sort(bySortThenName);
    return {
        tableLogicalName: (row.ce_tablelogicalname ?? "").toLowerCase(),
        name: row.ce_name || row.ce_tablelogicalname || "",
        contactLookupColumn: (row.ce_contactlookupcolumn ?? "").toLowerCase(),
        maxUploadMb: row.ce_maxuploadmb ?? undefined,
        categories,
    };
}

interface SeedFile {
    categories: { value: string; name: string; folderPath: string; color?: string; sortOrder?: number }[];
    tables: { tableLogicalName: string; name?: string; contactLookupColumn: string; maxUploadMb?: number; categories: string[] }[];
}

/** Builds table configs from a seed.json-shaped object (local development and tests). */
export function fromSeed(seed: SeedFile): Map<string, TableConfig> {
    const cats = new Map(seed.categories.map(c => [c.value, { ...c, sortOrder: c.sortOrder ?? 1000 }]));
    return new Map(seed.tables.map(t => [t.tableLogicalName.toLowerCase(), {
        tableLogicalName: t.tableLogicalName.toLowerCase(),
        name: t.name ?? t.tableLogicalName,
        contactLookupColumn: t.contactLookupColumn.toLowerCase(),
        maxUploadMb: t.maxUploadMb,
        categories: t.categories.map(v => cats.get(v)).filter((c): c is CategoryDef => !!c).sort(bySortThenName),
    }]));
}

export class StaticConfigProvider implements ConfigProvider {
    constructor(private tables: Map<string, TableConfig>) {}
    async getTable(name: string) { return this.tables.get(name.toLowerCase()); }
}

/** Reads configuration from Dataverse as the backend's application user, cached per table. */
export class DataverseConfigProvider implements ConfigProvider {
    private cache = new Map<string, { at: number; value: Promise<TableConfig | undefined> }>();
    constructor(private ttlMs = config.configCacheSeconds * 1000, private now = () => Date.now()) {}

    getTable(name: string): Promise<TableConfig | undefined> {
        const key = name.toLowerCase();
        const hit = this.cache.get(key);
        if (hit && this.now() - hit.at < this.ttlMs) return hit.value;
        const value = this.load(key).catch(e => { this.cache.delete(key); throw e; }); // don't cache failures
        this.cache.set(key, { at: this.now(), value });
        return value;
    }

    private async load(table: string): Promise<TableConfig | undefined> {
        if (!/^[a-z][a-z0-9_]{0,127}$/.test(table)) return undefined;
        const r = await dataverse<{ value: DvConfigRow[] }>(
            `/ce_documentmanagerconfigs?$select=ce_name,ce_tablelogicalname,ce_contactlookupcolumn,ce_maxuploadmb`
            + `&$filter=ce_tablelogicalname eq '${table}' and statecode eq 0`
            + `&$expand=ce_documentmanagerconfig_documentcategory($select=ce_name,ce_value,ce_folderpath,ce_color,ce_sortorder,statecode)`);
        const row = r.value?.[0];
        return row ? fromDataverseRow(row) : undefined;
    }
}

let provider: ConfigProvider | null = null;

export function configProvider(): ConfigProvider {
    if (!provider) {
        provider = config.docConfigFile
            ? new StaticConfigProvider(fromSeed(JSON.parse(readFileSync(config.docConfigFile, "utf8")) as SeedFile))
            : new DataverseConfigProvider();
    }
    return provider;
}

/** Tests only. */
export function setConfigProvider(p: ConfigProvider | null): void {
    provider = p;
}

export interface UploadPlan { table: TableConfig; category: CategoryDef; folderPath: string; maxBytes: number }

/**
 * Validates an upload against configuration: the table must be configured and active, the category
 * must be active and allowed on that table. Returns the folder and the effective size limit.
 */
export async function planUpload(tableLogicalName: string, categoryValue: string, globalMaxBytes: number): Promise<UploadPlan> {
    const table = await configProvider().getTable(tableLogicalName);
    if (!table) throw new ConfigError(`Document management is not configured for table "${tableLogicalName}"`);
    const category = table.categories.find(c => c.value === categoryValue);
    if (!category) throw new ConfigError(`Category "${categoryValue}" is not allowed on ${table.name}`);
    const tableMax = table.maxUploadMb ? table.maxUploadMb * 1024 * 1024 : Infinity;
    return { table, category, folderPath: safeFolderPath(category.folderPath), maxBytes: Math.min(globalMaxBytes, tableMax) };
}
