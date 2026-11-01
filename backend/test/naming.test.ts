import { test } from "node:test";
import assert from "node:assert/strict";
import { contactFolderName, encodePath, isOfficeFile, isValidEntityName, normalizeGuid, sanitizeName } from "../src/lib/naming";
import { toDto } from "../src/lib/sharepoint";

const GUID = "a1b2c3d4-0000-4000-8000-1234567890ab";

test("normalizeGuid accepts braces/upper case and rejects junk", () => {
    assert.equal(normalizeGuid(`{${GUID.toUpperCase()}}`), GUID);
    assert.equal(normalizeGuid("not-a-guid"), null);
    assert.equal(normalizeGuid("' or 1 eq 1"), null);
});

test("contact folder follows {FullName}_{GUID}", () => {
    assert.equal(contactFolderName("John Doe", GUID), `John Doe_${GUID}`);
    assert.equal(contactFolderName('Jo/hn: "D"', GUID), `Jo_hn_ _D__${GUID}`);
    assert.equal(contactFolderName("", GUID), `Contact_${GUID}`);
});

test("sanitizeName strips illegal characters and trailing dots", () => {
    assert.equal(sanitizeName("a<b>c?.pdf"), "a_b_c_.pdf");
    assert.equal(sanitizeName("report. "), "report");
});

test("encodePath encodes each segment", () => {
    assert.equal(encodePath("/John Doe_x/Offer Letters/a#1.pdf"), "John%20Doe_x/Offer%20Letters/a%231.pdf");
});

test("entity name validation", () => {
    assert.ok(isValidEntityName("opportunity"));
    assert.ok(isValidEntityName("ce_advisor_onboarding"));
    assert.ok(!isValidEntityName("Opportunity'"));
});

test("office detection", () => {
    assert.ok(isOfficeFile("x.DOCX"));
    assert.ok(!isOfficeFile("x.pdf"));
});

test("toDto maps fields and defaults the source to contact root", () => {
    const dto = toDto(
        { id: "1", name: "Form_U4.PDF", webUrl: "u", size: 10, createdDateTime: "2022-03-11T00:00:00Z" },
        { DocCategory: "FinCEN_U4", SourceRecordId: GUID.toUpperCase(), ContactGuid: GUID },
    );
    assert.equal(dto.sourceEntity, "contact");
    assert.equal(dto.sourceRecordId, GUID);
    assert.equal(dto.fileType, "pdf");
    assert.equal(dto.category, "FinCEN_U4");
});

import { retryDelayMs } from "../src/lib/graph";

test("retryDelayMs handles seconds, HTTP-dates and junk", () => {
    assert.equal(retryDelayMs("5", 0), 5000);
    assert.equal(retryDelayMs("120", 0), 30000); // capped
    const now = Date.parse("2026-11-01T00:00:00Z");
    assert.equal(retryDelayMs("Sun, 01 Nov 2026 00:00:07 GMT", 0, now), 7000);
    assert.equal(retryDelayMs("garbage", 2), 4000); // exponential fallback, never NaN
    assert.equal(retryDelayMs(null, 1), 2000);
});
