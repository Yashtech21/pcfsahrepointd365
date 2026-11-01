import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import {
    ConfigError, DataverseConfigProvider, StaticConfigProvider, fromDataverseRow, fromSeed, planUpload,
    safeFolderPath, setConfigProvider,
} from "../src/lib/tableConfig";

const seed = JSON.parse(readFileSync(join(__dirname, "..", "..", "..", "dataverse", "seed.json"), "utf8"));

test("seed.json (shared with the provisioning script) maps to per-table configs", () => {
    const tables = fromSeed(seed);
    const opp = tables.get("opportunity")!;
    assert.equal(opp.contactLookupColumn, "parentcontactid");
    assert.deepEqual(opp.categories.map(c => c.value), ["Offer Letter", "Commissions", "Misc"]);
    assert.equal(tables.get("ce_advisoronboarding")!.contactLookupColumn, "ce_recruitid");
});

test("Dataverse rows: inactive categories dropped, sorted, lookup column kept", () => {
    const cfg = fromDataverseRow({
        ce_name: "Advisor Onboarding", ce_tablelogicalname: "CE_AdvisorOnboarding", ce_contactlookupcolumn: "ce_RecruitId", ce_maxuploadmb: 25,
        ce_documentmanagerconfig_documentcategory: [
            { ce_name: "Misc", ce_value: "Misc", ce_folderpath: "Miscellaneous", ce_sortorder: 90, statecode: 0 },
            { ce_name: "Old", ce_value: "Old", ce_folderpath: "Old", ce_sortorder: 1, statecode: 1 },
            { ce_name: "Payslip", ce_value: "Payslip", ce_folderpath: "Onboarding Documents/Payslips", ce_sortorder: 30, statecode: 0 },
        ],
    });
    assert.equal(cfg.tableLogicalName, "ce_advisoronboarding");
    assert.equal(cfg.contactLookupColumn, "ce_recruitid");
    assert.equal(cfg.maxUploadMb, 25);
    assert.deepEqual(cfg.categories.map(c => c.value), ["Payslip", "Misc"]);
});

test("folder paths from configuration can't escape the contact folder", () => {
    assert.equal(safeFolderPath(" Onboarding Documents / Payslips "), "Onboarding Documents/Payslips");
    assert.equal(safeFolderPath("A\\B"), "A/B");
    assert.equal(safeFolderPath("Bad:Name"), "Bad_Name");
    assert.throws(() => safeFolderPath("../x"), ConfigError);
    assert.throws(() => safeFolderPath(""), ConfigError);
    assert.throws(() => safeFolderPath("a/b/c/d/e/f"), ConfigError);
});

test("planUpload enforces configured tables, allowed categories and per-table size", async () => {
    setConfigProvider(new StaticConfigProvider(fromSeed({
        categories: seed.categories,
        tables: [{ tableLogicalName: "opportunity", contactLookupColumn: "parentcontactid", maxUploadMb: 10, categories: ["Offer Letter"] }],
    })));
    try {
        const plan = await planUpload("opportunity", "Offer Letter", 100 * 1024 * 1024);
        assert.equal(plan.folderPath, "Offer Letters");
        assert.equal(plan.maxBytes, 10 * 1024 * 1024);
        assert.equal((await planUpload("opportunity", "Offer Letter", 5)).maxBytes, 5, "global limit still caps");
        await assert.rejects(planUpload("account", "Offer Letter", 1), /not configured for table "account"/);
        await assert.rejects(planUpload("opportunity", "Payslip", 1), /not allowed on opportunity/);
    } finally {
        setConfigProvider(null);
    }
});

test("Dataverse provider caches per table and doesn't cache failures", async () => {
    let t = 0;
    const p = new DataverseConfigProvider(1000, () => t);
    let calls = 0;
    (p as unknown as { load: (n: string) => Promise<unknown> }).load = async () => {
        calls++;
        if (calls === 1) throw new Error("boom");
        return { tableLogicalName: "opportunity" };
    };
    await assert.rejects(p.getTable("opportunity"));
    await p.getTable("opportunity");
    await p.getTable("Opportunity");
    assert.equal(calls, 2, "failure not cached; second success cached (case-insensitive)");
    t = 1500;
    await p.getTable("opportunity");
    assert.equal(calls, 3, "expired after TTL");
});
