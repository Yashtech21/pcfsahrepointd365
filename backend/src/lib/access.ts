import { Caller } from "./auth";
import { config } from "./config";
import { dataverseAsUser, GraphError } from "./graph";
import { DocumentDto } from "./sharepoint";
import { TableConfig } from "./tableConfig";

/**
 * Per-record security. Every check runs in Dataverse **as the signed-in user** (On-Behalf-Of), so the
 * user's own security roles, business units, sharing and teams decide the outcome, exactly as in D365.
 *
 *   list      → read the contact; documents whose source record the user can't read are hidden
 *   upload    → read the contact; the source record must point to that contact (its configured
 *               Contact Lookup Column); CHANGE_ACCESS ("write" by default) on the source record
 *   preview / download → read the contact and the document's source record
 *   delete    → read the contact; CHANGE_ACCESS on the source record (or on the contact for root documents)
 */
export type AccessLevel = "read" | "write";

/** The user may not do this (no access, or the record isn't visible to them). → 403 */
export class AccessDeniedError extends Error {}
/** The source record isn't linked to the contact the client sent (stale form or forged request). → 409 */
export class RecordContactMismatchError extends Error {}

/** Minimal Dataverse transport, so the checker can be tested without Dataverse. */
export type DataverseGet = <T>(path: string) => Promise<T>;

export interface AccessChecker {
    assertContact(contactId: string, level: AccessLevel): Promise<void>;
    assertRecord(table: string, recordId: string, level: AccessLevel): Promise<void>;
    /** Reads the record as the user and checks its contact lookup equals `contactId`. */
    assertRecordLinkedToContact(table: TableConfig, recordId: string, contactId: string): Promise<void>;
    /** Removes documents whose source record the user can't read. */
    filterReadable(docs: DocumentDto[]): Promise<DocumentDto[]>;
}

interface TableMeta { entitySet: string; primaryId: string }

const LOGICAL_NAME = /^[a-z][a-z0-9_]{0,127}$/;
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;
const IN_CHUNK = 50;

/** Table metadata is the same for every user: cache it for the life of the worker. */
const metaCache = new Map<string, Promise<TableMeta | null>>();

/** Dataverse answers 404 (not visible) or 403 (no privilege) when the user can't access a record. */
const isDenied = (e: unknown) => e instanceof GraphError && (e.status === 403 || e.status === 404);

export class DataverseAccessChecker implements AccessChecker {
    private decisions = new Map<string, Promise<void>>();   // per request: don't ask twice
    private whoAmI: Promise<string> | null = null;

    constructor(private get: DataverseGet, private changeLevel: AccessLevel = config.changeAccess) {}

    private meta(table: string): Promise<TableMeta | null> {
        if (!LOGICAL_NAME.test(table)) return Promise.resolve(null);
        if (table === "contact") return Promise.resolve({ entitySet: "contacts", primaryId: "contactid" });
        let p = metaCache.get(table);
        if (!p) {
            p = this.get<{ EntitySetName?: string; PrimaryIdAttribute?: string }>(
                `/EntityDefinitions(LogicalName='${table}')?$select=EntitySetName,PrimaryIdAttribute`)
                .then(m => (m.EntitySetName && m.PrimaryIdAttribute ? { entitySet: m.EntitySetName, primaryId: m.PrimaryIdAttribute } : null))
                .catch(e => {
                    metaCache.delete(table);
                    if (isDenied(e)) return null;
                    throw e;
                });
            metaCache.set(table, p);
        }
        return p;
    }

    private async userId(): Promise<string> {
        this.whoAmI ??= this.get<{ UserId: string }>("/WhoAmI").then(r => r.UserId);
        return this.whoAmI;
    }

    private check(table: string, id: string, level: AccessLevel, what: string): Promise<void> {
        const key = `${table}|${id}|${level}`;
        let p = this.decisions.get(key);
        if (!p) {
            p = this.evaluate(table, id, level, what);
            this.decisions.set(key, p);
        }
        return p;
    }

    private async evaluate(table: string, id: string, level: AccessLevel, what: string): Promise<void> {
        if (!GUID.test(id)) throw new AccessDeniedError(`Invalid ${what} id`);
        const meta = await this.meta(table);
        if (!meta) throw new AccessDeniedError(`You don't have access to ${what} records (${table})`);
        try {
            // Read: the record must be visible to the user.
            await this.get(`/${meta.entitySet}(${id})?$select=${meta.primaryId}`);
        } catch (e) {
            if (isDenied(e)) throw new AccessDeniedError(`You don't have access to this ${what}`);
            throw e;
        }
        if (level === "write") {
            const uid = await this.userId();
            const target = encodeURIComponent(JSON.stringify({ "@odata.id": `${meta.entitySet}(${id})` }));
            const r = await this.get<{ AccessRights?: string }>(
                `/systemusers(${uid})/Microsoft.Dynamics.CRM.RetrievePrincipalAccess(Target=@tid)?@tid=${target}`);
            if (!(r.AccessRights ?? "").split(/,\s*/).includes("WriteAccess")) {
                throw new AccessDeniedError(`You can view this ${what} but don't have permission to change its documents`);
            }
        }
    }

    assertContact(contactId: string, level: AccessLevel) {
        return this.check("contact", contactId, level, "contact");
    }

    assertRecord(table: string, recordId: string, level: AccessLevel) {
        return this.check(table, recordId, level, "record");
    }

    async assertRecordLinkedToContact(table: TableConfig, recordId: string, contactId: string): Promise<void> {
        const meta = await this.meta(table.tableLogicalName);
        if (!meta) throw new AccessDeniedError(`You don't have access to ${table.name} records`);
        const col = table.contactLookupColumn;
        if (!LOGICAL_NAME.test(col)) throw new AccessDeniedError(`The configured contact lookup column for ${table.name} is invalid`);
        let row: Record<string, unknown>;
        try {
            row = await this.get(`/${meta.entitySet}(${recordId})?$select=_${col}_value`);
        } catch (e) {
            if (isDenied(e)) throw new AccessDeniedError(`You don't have access to this ${table.name} record`);
            if (e instanceof GraphError && e.status === 400) {
                throw new AccessDeniedError(`"${col}" isn't a lookup column on ${table.tableLogicalName}; fix the Contact Lookup Column in its Document Manager Configuration`);
            }
            throw e;
        }
        const linked = String(row[`_${col}_value`] ?? "").toLowerCase();
        if (linked !== contactId) {
            throw new RecordContactMismatchError(linked
                ? `This ${table.name} is linked to a different contact. Save or refresh the form and try again.`
                : `This ${table.name} has no contact selected. Choose one, save, and try again.`);
        }
    }

    async filterReadable(docs: DocumentDto[]): Promise<DocumentDto[]> {
        // Group by source table; contact-root documents (and documents without a source) follow the contact check.
        const byTable = new Map<string, Set<string>>();
        for (const d of docs) {
            if (!d.sourceRecordId || d.sourceEntity === "contact" || !GUID.test(d.sourceRecordId)) continue;
            if (!byTable.has(d.sourceEntity)) byTable.set(d.sourceEntity, new Set());
            byTable.get(d.sourceEntity)!.add(d.sourceRecordId);
        }
        const readable = new Map<string, Set<string>>();
        for (const [table, ids] of byTable) {
            const meta = await this.meta(table);
            const ok = new Set<string>();
            readable.set(table, ok);
            if (!meta) continue;                       // table not readable/unknown → its documents are hidden
            const all = [...ids];
            for (let i = 0; i < all.length; i += IN_CHUNK) {
                const chunk = all.slice(i, i + IN_CHUNK);
                const filter = `Microsoft.Dynamics.CRM.In(PropertyName='${meta.primaryId}',PropertyValues=[${chunk.map(x => `'${x}'`).join(",")}])`;
                try {
                    const r = await this.get<{ value: Record<string, string>[] }>(`/${meta.entitySet}?$select=${meta.primaryId}&$filter=${filter}`);
                    r.value.forEach(row => ok.add(String(row[meta.primaryId]).toLowerCase()));
                } catch (e) {
                    if (!isDenied(e)) throw e;          // no read privilege on the table → none readable
                }
            }
        }
        return docs.filter(d => {
            if (!d.sourceRecordId || d.sourceEntity === "contact" || !GUID.test(d.sourceRecordId)) return true;
            return readable.get(d.sourceEntity)?.has(d.sourceRecordId) ?? false;
        });
    }
}

/** Local development only (REQUIRE_AUTH=false + ACCESS_CHECKS=off): allows everything. */
export class AllowAllAccessChecker implements AccessChecker {
    async assertContact() { /* allow */ }
    async assertRecord() { /* allow */ }
    async assertRecordLinkedToContact() { /* allow */ }
    async filterReadable(docs: DocumentDto[]) { return docs; }
}

/** The checker for one request. Access checks can only be turned off together with authentication. */
export function accessFor(caller: Caller): AccessChecker {
    if (!config.requireAuth) {
        if (!config.accessChecks) return new AllowAllAccessChecker();
        throw new Error("ACCESS_CHECKS needs a signed-in user: set ACCESS_CHECKS=off when REQUIRE_AUTH=false (local development only)");
    }
    return new DataverseAccessChecker(<T>(path: string) => dataverseAsUser<T>(caller.token, path));
}

/** Tests only. */
export function clearAccessMetadataCache(): void {
    metaCache.clear();
}
