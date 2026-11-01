import { test } from "node:test";
import assert from "node:assert/strict";
import { assertFileAllowed } from "../src/lib/fileRules";
import { RejectedFileError } from "../src/lib/errors";
import { retryDecision } from "../src/lib/graph";
import { sanitizeName } from "../src/lib/naming";
import { assertTrustedUrl } from "../src/lib/sharepoint";

const bytes = (s: string) => new Uint8Array(Buffer.from(s, "binary"));
const rejects = (name: string, content: Uint8Array) =>
    assert.throws(() => assertFileAllowed(name, content), RejectedFileError, name);

test("blocked extensions, including double extensions", () => {
    rejects("payroll.exe", bytes("hello"));
    rejects("invoice.pdf.exe", bytes("%PDF-1.7"));
    rejects("macro.xlsm", bytes("PK\x03\x04"));
});

test("executables are rejected whatever their name", () => {
    rejects("offer.pdf", bytes("MZ\x90\x00"));
    rejects("notes.txt", bytes("#!/bin/sh\nrm -rf /"));
});

test("content must match known extensions", () => {
    rejects("offer.pdf", bytes("PK\x03\x04 actually a zip"));
    rejects("form.docx", bytes("%PDF-1.4"));
    assert.doesNotThrow(() => assertFileAllowed("offer.pdf", bytes("%PDF-1.7 ...")));
    assert.doesNotThrow(() => assertFileAllowed("form.docx", bytes("PK\x03\x04....")));
    assert.doesNotThrow(() => assertFileAllowed("legacy.doc", new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0])));
    assert.doesNotThrow(() => assertFileAllowed("notes.txt", bytes("plain text")));
});

test("only safe failures are retried for non-idempotent requests", () => {
    assert.equal(retryDecision(429, false, 0), "retry");
    assert.equal(retryDecision(503, false, 0), "retry");
    assert.equal(retryDecision("timeout", false, 0), "fail"); // upload PUT must not duplicate
    assert.equal(retryDecision(504, false, 0), "fail");
    assert.equal(retryDecision("timeout", true, 0), "retry");
    assert.equal(retryDecision("network", true, 0), "retry");
    assert.equal(retryDecision(429, true, 4), "fail");          // attempts exhausted
    assert.equal(retryDecision(400, true, 0), "fail");
});

test("sanitizeName handles SharePoint-reserved patterns and keeps the extension", () => {
    assert.equal(sanitizeName("~$draft.docx"), "draft.docx");
    assert.equal(sanitizeName("a_vti_b.pdf"), "a_vti-b.pdf");
    assert.equal(sanitizeName("CON.pdf"), "_CON.pdf");
    assert.equal(sanitizeName("..hidden.pdf"), "hidden.pdf");
    const long = sanitizeName("x".repeat(300) + ".pdf");
    assert.equal(long.length, 200);
    assert.ok(long.endsWith(".pdf"));
});

test("only https Microsoft/SharePoint URLs are passed to the browser", () => {
    assert.ok(assertTrustedUrl("https://contoso.sharepoint.com/:w:/r/x?action=embedview"));
    assert.throws(() => assertTrustedUrl("http://contoso.sharepoint.com/x"));
    assert.throws(() => assertTrustedUrl("https://evil.example.com/x"));
    assert.throws(() => assertTrustedUrl("https://contoso.sharepoint.com.evil.io/x"));
    assert.throws(() => assertTrustedUrl("javascript:alert(1)"));
});
