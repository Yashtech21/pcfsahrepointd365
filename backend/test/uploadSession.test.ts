import { test } from "node:test";
import assert from "node:assert/strict";
import { uploadToSession } from "../src/lib/uploadSession";
import { GraphError } from "../src/lib/graph";

const URL_ = "https://tenant.sharepoint.com/upload/session";
const noSleep = async () => undefined;

type Step = (init: RequestInit & { method?: string }) => Response | Promise<Response> | "timeout" | "network";

/** Fake upload endpoint: tracks which bytes the "server" has, scripted failures per call index. */
function fakeServer(total: number, script: Record<number, Step> = {}) {
    let received = 0;
    const calls: { method: string; range?: string }[] = [];
    const fetchImpl = (async (_url: string, init: RequestInit = {}) => {
        const method = (init.method ?? "GET").toUpperCase();
        const range = (init.headers as Record<string, string> | undefined)?.["Content-Range"];
        const idx = calls.length;
        calls.push({ method, range });
        const scripted = script[idx];
        if (scripted) {
            const r = await scripted(init);
            if (r === "timeout") throw Object.assign(new Error("timed out"), { name: "TimeoutError" });
            if (r === "network") throw new TypeError("fetch failed");
            if (method === "PUT" && r.status === 202) received = Number(range!.split(/[ -/]/)[2]) + 1; // landed
            return r;
        }
        if (method === "GET") return Response.json({ nextExpectedRanges: [`${received}-`] });
        if (method === "DELETE") return new Response(null, { status: 204 });
        const [start, end] = range!.replace("bytes ", "").split("/")[0].split("-").map(Number);
        if (start !== received) return new Response("bad range", { status: 416 });
        received = end + 1;
        return received === total
            ? Response.json({ id: "item1", name: "f.pdf", size: total }, { status: 201 })
            : Response.json({ nextExpectedRanges: [`${received}-`] }, { status: 202 });
    }) as unknown as typeof fetch;
    return { fetchImpl, calls, received: () => received };
}

const data = (n: number) => new Uint8Array(n).fill(7);

test("uploads in chunks and returns the item", async () => {
    const srv = fakeServer(25);
    const item = await uploadToSession<{ id: string }>(URL_, data(25), { chunkSize: 10, fetchImpl: srv.fetchImpl, sleepImpl: noSleep });
    assert.equal(item.id, "item1");
    assert.deepEqual(srv.calls.map(c => c.range), ["bytes 0-9/25", "bytes 10-19/25", "bytes 20-24/25"]);
});

test("throttled chunk (429) is resent after waiting", async () => {
    const waits: number[] = [];
    const srv = fakeServer(20, { 1: () => new Response("", { status: 429, headers: { "Retry-After": "3" } }) });
    await uploadToSession(URL_, data(20), { chunkSize: 10, fetchImpl: srv.fetchImpl, sleepImpl: async ms => { waits.push(ms); } });
    assert.deepEqual(waits, [3000]);
    assert.deepEqual(srv.calls.map(c => c.range).filter(Boolean), ["bytes 0-9/20", "bytes 10-19/20", "bytes 10-19/20"]);
});

test("timeout after the chunk actually landed resumes from nextExpectedRanges (no double send)", async () => {
    // Call 1 (second chunk) lands on the server but the response is lost.
    const srv = fakeServer(30, {});
    let landed = false;
    const inner = srv.fetchImpl;
    const fetchImpl = (async (u: string, init: RequestInit = {}) => {
        const range = (init.headers as Record<string, string> | undefined)?.["Content-Range"];
        if (range === "bytes 10-19/30" && !landed) { landed = true; await inner(u, init); throw Object.assign(new Error("t"), { name: "TimeoutError" }); }
        return inner(u, init);
    }) as unknown as typeof fetch;
    const item = await uploadToSession<{ id: string }>(URL_, data(30), { chunkSize: 10, fetchImpl, sleepImpl: noSleep });
    assert.equal(item.id, "item1");
    const puts = srv.calls.filter(c => c.method === "PUT").map(c => c.range);
    assert.deepEqual(puts, ["bytes 0-9/30", "bytes 10-19/30", "bytes 20-29/30"]); // 10-19 sent once
    assert.ok(srv.calls.some(c => c.method === "GET"), "asked the session where to resume");
});

test("repeated timeouts give up with 504 and cancel the session", async () => {
    const script: Record<number, Step> = {};
    for (let i = 0; i < 40; i++) script[i] = (init) => ((init.method ?? "GET") === "PUT" ? "timeout" : Response.json({ nextExpectedRanges: ["0-"] }));
    const srv = fakeServer(10, script);
    await assert.rejects(
        uploadToSession(URL_, data(10), { chunkSize: 10, fetchImpl: srv.fetchImpl, sleepImpl: noSleep }),
        (e: unknown) => e instanceof GraphError && e.status === 504,
    );
    assert.equal(srv.calls.at(-1)!.method, "DELETE");
});

test("expired session (404) fails fast and cancels", async () => {
    const srv = fakeServer(10, { 0: () => new Response("gone", { status: 404 }) });
    await assert.rejects(uploadToSession(URL_, data(10), { chunkSize: 10, fetchImpl: srv.fetchImpl, sleepImpl: noSleep }),
        (e: unknown) => e instanceof GraphError && e.status === 404);
    assert.deepEqual(srv.calls.map(c => c.method), ["PUT", "DELETE"]);
});

test("cancel request carries a timeout signal", async () => {
    let cancelSignal: AbortSignal | undefined;
    const srv = fakeServer(10, {
        0: () => new Response("nope", { status: 400 }),
        1: (init) => { cancelSignal = init.signal ?? undefined; return new Response(null, { status: 204 }); },
    });
    await assert.rejects(uploadToSession(URL_, data(10), { chunkSize: 10, fetchImpl: srv.fetchImpl, sleepImpl: noSleep }));
    assert.ok(cancelSignal, "DELETE must have an AbortSignal");
});
