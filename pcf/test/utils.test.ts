/**
 * Unit tests for browser-independent helpers in the control. Run: npm test (pcf/).
 * Browser globals are stubbed per test.
 */
import { test, afterEach } from "node:test";
import assert from "node:assert/strict";
import { copyText, isMacPlatform, formatSize, normalizeGuid, prettifyEntityName } from "../CategorizedDocManager/utils";

const g = globalThis as unknown as { window?: unknown; navigator?: unknown };
const realNavigator = Object.getOwnPropertyDescriptor(globalThis, "navigator");

function stub(win: object, nav: object) {
    g.window = win;
    Object.defineProperty(globalThis, "navigator", { value: nav, configurable: true, writable: true });
}

afterEach(() => {
    delete g.window;
    if (realNavigator) Object.defineProperty(globalThis, "navigator", realNavigator);
    else delete g.navigator;
});

test("copyText writes to the clipboard in a secure context", async () => {
    let written = "";
    stub({ isSecureContext: true }, { clipboard: { writeText: async (t: string) => { written = t; } } });
    assert.equal(await copyText("https://contoso.sharepoint.com/a.pdf"), true);
    assert.equal(written, "https://contoso.sharepoint.com/a.pdf");
});

test("copyText returns false (→ fallback dialog) when the browser refuses", async () => {
    stub({ isSecureContext: true }, { clipboard: { writeText: async () => { throw new Error("NotAllowedError"); } } });
    assert.equal(await copyText("x"), false);
});

test("copyText returns false without a secure context or Clipboard API", async () => {
    stub({ isSecureContext: false }, { clipboard: { writeText: async () => undefined } });
    assert.equal(await copyText("x"), false);
    stub({ isSecureContext: true }, {});
    assert.equal(await copyText("x"), false);
});

test("isMacPlatform prefers UA Client Hints, falls back to userAgent", () => {
    assert.equal(isMacPlatform({ userAgentData: { platform: "macOS" } }), true);
    assert.equal(isMacPlatform({ userAgentData: { platform: "Windows" }, userAgent: "Macintosh" }), false);
    assert.equal(isMacPlatform({ userAgent: "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_0)" }), true);
    assert.equal(isMacPlatform({ userAgent: "Mozilla/5.0 (iPad; CPU OS 17_0)" }), true);
    assert.equal(isMacPlatform({ userAgent: "Mozilla/5.0 (Windows NT 10.0)" }), false);
    assert.equal(isMacPlatform({}), false);
});

test("small formatters", () => {
    assert.equal(normalizeGuid("{AB-CD}"), "ab-cd");
    assert.equal(prettifyEntityName("advisor_onboarding"), "Advisor Onboarding");
    assert.equal(formatSize(0), "");
    assert.equal(formatSize(1536), "2 KB");
});
