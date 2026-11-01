import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import {
    validateSeed, tableDefinition, relationshipDefinition, TABLES, RELATIONSHIP, parseArgs,
    ensureSchema, seed, setEnvValues, ensureAppUser, odataString, main,
} from "./provision.mjs";

const seedData = JSON.parse(readFileSync(new URL("./seed.json", import.meta.url), "utf8"));
const quiet = () => {};

/** Fake client: `lists` maps a path prefix to rows; every write is recorded. */
function fakeClient({ exists = () => true, lists = {} } = {}) {
    const writes = [];
    const pick = (p) => Object.entries(lists).find(([k]) => p.startsWith(k))?.[1] ?? [];
    return {
        writes, base: "https://x.crm.dynamics.com/api/data/v9.2", offline: false,
        exists: async (p) => exists(p),
        list: async (p) => pick(p),
        get: async (p) => (p.includes("EntitySetName") ? { EntitySetName: p.includes("ce_documentcategory'") ? "ce_documentcategories" : "ce_documentmanagerconfigs" } : { value: pick(p) }),
        post: async (p, b) => { writes.push(["POST", p, b]); return { id: "11111111-1111-1111-1111-111111111111" }; },
        patch: async (p, b) => { writes.push(["PATCH", p, b]); return {}; },
    };
}

test("shipped seed.json is valid", () => assert.deepEqual(validateSeed(seedData), []));

test("seed validation catches typos before anything is written", () => {
    const bad = {
        categories: [{ value: "A", name: "A", folderPath: "x/../y" }, { value: "A", name: "A2", folderPath: "z", color: "red" }],
        tables: [{ tableLogicalName: "Opportunity", contactLookupColumn: "parentcontactid", categories: ["B"] }],
    };
    const errors = validateSeed(bad).join("\n");
    assert.match(errors, /invalid folderPath/);
    assert.match(errors, /duplicate category value "A"/);
    assert.match(errors, /colour/);
    assert.match(errors, /not a valid logical name/);
    assert.match(errors, /unknown category "B"/);
});

test("table definitions are organization-owned with a primary name column", () => {
    for (const t of TABLES) {
        const d = tableDefinition(t);
        assert.equal(d.OwnershipType, "OrganizationOwned");
        assert.equal(d.Attributes.length, 1);
        assert.equal(d.Attributes[0].IsPrimaryName, true);
        assert.equal(d.SchemaName.startsWith("ce_"), true);
    }
    const contactCol = TABLES[1].attributes.find(a => a.SchemaName === "ce_ContactLookupColumn");
    assert.equal(contactCol.RequiredLevel.Value, "ApplicationRequired", "contact lookup column must be required");
});

test("N:N relationship links configuration to categories", () => {
    const r = relationshipDefinition();
    assert.equal(r.SchemaName, RELATIONSHIP);
    assert.equal(r.Entity1LogicalName, "ce_documentmanagerconfig");
    assert.equal(r.Entity2LogicalName, "ce_documentcategory");
});

test("re-running schema when everything exists writes nothing", async () => {
    const c = fakeClient();
    await ensureSchema(c, { log: quiet });
    assert.deepEqual(c.writes, []);
});

test("schema creates only what is missing", async () => {
    const c = fakeClient({ exists: (p) => !p.includes("ce_maxuploadmb") });
    await ensureSchema(c, { log: quiet });
    assert.equal(c.writes.length, 1);
    assert.equal(c.writes[0][2].SchemaName, "ce_MaxUploadMb");
});

test("seed is create-only and only links missing categories", async () => {
    const lists = {
        "/ce_documentcategories": [{ ce_documentcategoryid: "c1" }],
        "/ce_documentmanagerconfigs": [{ ce_documentmanagerconfigid: "t1", [RELATIONSHIP]: [{ ce_value: "Offer Letter" }, { ce_value: "Commissions" }, { ce_value: "Misc" }, { ce_value: "Payslip" }, { ce_value: "Resignation Letter" }] }],
    };
    const c = fakeClient({ lists });
    await seed(c, seedData, { log: quiet });
    // Everything exists; only FinCEN_U4 isn't linked to the second table yet.
    assert.equal(c.writes.length, 1);
    assert.match(c.writes[0][1], /\/ce_documentmanagerconfigs\(t1\)\/ce_documentmanagerconfig_documentcategory\/\$ref$/);
});

test("set-env updates an existing value, creates a missing one", async () => {
    const lists = { "/environmentvariabledefinitions": [{ environmentvariabledefinitionid: "d1", environmentvariabledefinition_environmentvariablevalue: [{ environmentvariablevalueid: "v1" }] }] };
    const c = fakeClient({ lists });
    await setEnvValues(c, { ce_DocManagerTenantId: "t" }, { log: quiet });
    assert.deepEqual(c.writes, [["PATCH", "/environmentvariablevalues(v1)", { value: "t" }]]);

    const c2 = fakeClient({ lists: { "/environmentvariabledefinitions": [{ environmentvariabledefinitionid: "d1" }] } });
    await setEnvValues(c2, { ce_DocManagerTenantId: "t" }, { log: quiet });
    assert.equal(c2.writes[0][1], "/environmentvariablevalues");
    assert.equal(c2.writes[0][2]["EnvironmentVariableDefinitionId@odata.bind"], "/environmentvariabledefinitions(d1)");
});

test("app-user rejects a non-GUID app id", async () => {
    await assert.rejects(ensureAppUser(fakeClient(), "not-a-guid", { log: quiet }), /GUID/);
});

test("arguments and OData quoting", () => {
    assert.deepEqual(parseArgs(["set-env", "--org", "https://a.crm.dynamics.com", "--dry-run"]),
        { command: "set-env", org: "https://a.crm.dynamics.com", "dry-run": true });
    assert.throws(() => parseArgs(["all", "--org"]), /needs a value/);
    assert.equal(odataString("O'Brien"), "'O''Brien'");
});

test("org URL is validated", async () => {
    await assert.rejects(main(["all", "--org", "https://evil.example.com"], { log: quiet }), /crm\.dynamics\.com/);
});
