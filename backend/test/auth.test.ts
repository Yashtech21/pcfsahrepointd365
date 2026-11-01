import { test } from "node:test";
import assert from "node:assert/strict";
import { HttpRequest } from "@azure/functions";
import { AuthError, getCaller } from "../src/lib/auth";

process.env.GRAPH_TENANT_ID = "11111111-1111-4111-8111-111111111111";
process.env.GRAPH_CLIENT_ID = "22222222-2222-4222-8222-222222222222";

const principal = Buffer.from(JSON.stringify({
    claims: [{ typ: "http://schemas.microsoft.com/identity/claims/tenantid", val: process.env.GRAPH_TENANT_ID }],
})).toString("base64");

test("a forged X-MS-CLIENT-PRINCIPAL without a bearer token is rejected", async () => {
    const req = new HttpRequest({ url: "https://x/api/documents", method: "GET", headers: { "x-ms-client-principal": principal } });
    await assert.rejects(getCaller(req), (e: unknown) => e instanceof AuthError && e.status === 401);
});

test("a malformed bearer token is rejected", async () => {
    const req = new HttpRequest({ url: "https://x/api/documents", method: "GET", headers: { authorization: "Bearer not.a.jwt" } });
    await assert.rejects(getCaller(req), (e: unknown) => e instanceof AuthError && e.status === 401);
});
