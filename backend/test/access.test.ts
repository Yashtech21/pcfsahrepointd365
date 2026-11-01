import { test, beforeEach } from "node:test";
import assert from "node:assert/strict";
import {
    AccessDeniedError, AllowAllAccessChecker, DataverseAccessChecker, RecordContactMismatchError, accessFor,
    clearAccessMetadataCache,
} from "../src/lib/access";
import { GraphError } from "../src/lib/graph";
import { DocumentDto } from "../src/lib/sharepoint";
import { TableConfig } from "../src/lib/tableConfig";

const C = "11111111-1111-4111-8111-111111111111";
const OPP1 = "22222222-2222-4222-8222-222222222222";
const OPP2 = "33333333-3333-4333-8333-333333333333";
const USER = "44444444-4444-4444-8444-444444444444";

/** Fake Dataverse "as the user": readable records, write-access records, lookups. */
function fakeDv(opts: { readable?: string[]; writable?: string[]; lookups?: Record<string, string | null>; badColumn?: boolean } = {}) {
    const calls: string[] = [];
    const readable = new Set(opts.readable ?? []);
    const get = async <T>(path: string): Promise<T> => {
        calls.push(path);
        if (path.startsWith("/EntityDefinitions(LogicalName='opportunity')")) return { EntitySetName: "opportunities", PrimaryIdAttribute: "opportunityid" } as T;
        if (path.startsWith("/EntityDefinitions")) throw new GraphError(404, "no such table", undefined, "Dataverse");
        if (path === "/WhoAmI") return { UserId: USER } as T;
        if (path.includes("RetrievePrincipalAccess")) {
            const id = /\(([0-9a-f-]{36})\)%22/.exec(decodeURIComponent(path).replace(/"/g, "%22"))?.[1] ?? "";
            return { AccessRights: (opts.writable ?? []).includes(id) ? "ReadAccess, WriteAccess, AppendAccess" : "ReadAccess" } as T;
        }
        const list = /^\/opportunities\?\$select=opportunityid&\$filter=.*PropertyValues=\[(.*)\]\)$/.exec(path);
        if (list) {
            const ids = list[1].split(",").map(s => s.replace(/'/g, ""));
            return { value: ids.filter(i => readable.has(i)).map(i => ({ opportunityid: i })) } as T;
        }
        const one = /^\/(contacts|opportunities)\(([0-9a-f-]{36})\)\?\$select=(.*)$/.exec(path);
        if (one) {
            const [, , id, sel] = one;
            if (!readable.has(id)) throw new GraphError(404, "Does not exist", undefined, "Dataverse");
            if (sel.startsWith("_")) {
                if (opts.badColumn) throw new GraphError(400, "Could not find a property named '_x_value'", undefined, "Dataverse");
                return { [sel]: opts.lookups?.[id] ?? null } as T;
            }
            return {} as T;
        }
        throw new Error(`unexpected ${path}`);
    };
    return { get, calls };
}

const oppTable: TableConfig = { tableLogicalName: "opportunity", name: "Opportunity", contactLookupColumn: "parentcontactid", categories: [] };
const doc = (id: string, sourceEntity: string, sourceRecordId: string): DocumentDto => ({
    id, name: id, category: "", sourceEntity, sourceRecordId, contactId: C, uploadDate: "", webUrl: "", size: 0, fileType: "",
});

beforeEach(() => clearAccessMetadataCache());

test("contact read: visible passes, invisible is denied", async () => {
    await new DataverseAccessChecker(fakeDv({ readable: [C] }).get).assertContact(C, "read");
    await assert.rejects(new DataverseAccessChecker(fakeDv().get).assertContact(C, "read"), AccessDeniedError);
});

test("write needs WriteAccess from RetrievePrincipalAccess (as the user)", async () => {
    const dv = fakeDv({ readable: [OPP1, OPP2], writable: [OPP1] });
    const a = new DataverseAccessChecker(dv.get);
    await a.assertRecord("opportunity", OPP1, "write");
    await assert.rejects(a.assertRecord("opportunity", OPP2, "write"), /don't have permission to change/);
    assert.ok(dv.calls.includes("/WhoAmI"));
    assert.ok(dv.calls.some(c => c.startsWith(`/systemusers(${USER})/Microsoft.Dynamics.CRM.RetrievePrincipalAccess`)));
});

test("decisions are cached per request", async () => {
    const dv = fakeDv({ readable: [C] });
    const a = new DataverseAccessChecker(dv.get);
    await a.assertContact(C, "read");
    await a.assertContact(C, "read");
    assert.equal(dv.calls.filter(c => c.startsWith("/contacts(")).length, 1);
});

test("upload: record must point to the same contact via the configured lookup column", async () => {
    await new DataverseAccessChecker(fakeDv({ readable: [OPP1], lookups: { [OPP1]: C } }).get)
        .assertRecordLinkedToContact(oppTable, OPP1, C);

    const dv = fakeDv({ readable: [OPP1], lookups: { [OPP1]: "99999999-9999-4999-8999-999999999999" } });
    await assert.rejects(new DataverseAccessChecker(dv.get).assertRecordLinkedToContact(oppTable, OPP1, C), RecordContactMismatchError);
    assert.ok(dv.calls.some(c => c === `/opportunities(${OPP1})?$select=_parentcontactid_value`), "uses the configured column");

    await assert.rejects(new DataverseAccessChecker(fakeDv({ readable: [OPP1], lookups: {} }).get)
        .assertRecordLinkedToContact(oppTable, OPP1, C), /no contact selected/);
    await assert.rejects(new DataverseAccessChecker(fakeDv().get)
        .assertRecordLinkedToContact(oppTable, OPP1, C), AccessDeniedError);
    await assert.rejects(new DataverseAccessChecker(fakeDv({ readable: [OPP1], badColumn: true }).get)
        .assertRecordLinkedToContact(oppTable, OPP1, C), /isn't a lookup column/);
});

test("list: documents from unreadable records are hidden; contact-root documents stay", async () => {
    const a = new DataverseAccessChecker(fakeDv({ readable: [OPP1] }).get);
    const out = await a.filterReadable([
        doc("root", "contact", ""), doc("mine", "opportunity", OPP1), doc("theirs", "opportunity", OPP2),
        doc("unknown-table", "ce_gone", OPP1),
    ]);
    assert.deepEqual(out.map(d => d.id), ["root", "mine"]);
});

test("list: large lists are checked in chunks", async () => {
    const ids = Array.from({ length: 120 }, (_, i) => `aaaaaaaa-aaaa-4aaa-8aaa-${String(i).padStart(12, "0")}`);
    const dv = fakeDv({ readable: ids.filter((_, i) => i % 2 === 0) });
    const out = await new DataverseAccessChecker(dv.get).filterReadable(ids.map(i => doc(i, "opportunity", i)));
    assert.equal(out.length, 60);
    assert.equal(dv.calls.filter(c => c.startsWith("/opportunities?")).length, 3, "50 + 50 + 20");
});

test("checks can't be switched off in production", () => {
    const env = { ...process.env };
    try {
        Object.assign(process.env, { GRAPH_TENANT_ID: "t", GRAPH_CLIENT_ID: "c", GRAPH_CLIENT_SECRET: "s", DATAVERSE_URL: "https://x.crm.dynamics.com" });
        const caller = { userId: "u", name: "u", tenantId: "t", roles: [], token: "tok" };

        process.env.REQUIRE_AUTH = "true"; process.env.ACCESS_CHECKS = "off";
        assert.ok(accessFor(caller) instanceof DataverseAccessChecker, "ACCESS_CHECKS=off is ignored when auth is on");

        process.env.REQUIRE_AUTH = "false"; process.env.ACCESS_CHECKS = "on";
        assert.throws(() => accessFor(caller), /ACCESS_CHECKS=off/);

        process.env.ACCESS_CHECKS = "off";
        assert.ok(accessFor(caller) instanceof AllowAllAccessChecker);
    } finally {
        process.env = env;
    }
});
