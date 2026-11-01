import { createHash } from "node:crypto";
import { ClientSecretCredential, OnBehalfOfCredential } from "@azure/identity";
import { config } from "./config";

const GRAPH = "https://graph.microsoft.com/v1.0";
let credential: ClientSecretCredential | null = null;

const REQUEST_TIMEOUT_MS = Number(process.env.GRAPH_TIMEOUT_MS || "30000");
const MAX_RETRY_WAIT_MS = 30000;
const MAX_ATTEMPTS = 5;

/**
 * Retry-After may be delta-seconds or an HTTP-date. Anything unparsable falls back to
 * exponential backoff so a throttled call never retries instantly.
 */
export function retryDelayMs(header: string | null, attempt: number, now = Date.now()): number {
    const fallback = Math.min(2 ** attempt * 1000, MAX_RETRY_WAIT_MS);
    if (!header) return fallback;
    const seconds = Number(header);
    if (Number.isFinite(seconds) && seconds >= 0) return Math.min(seconds * 1000, MAX_RETRY_WAIT_MS);
    const date = Date.parse(header);
    if (Number.isFinite(date)) return Math.min(Math.max(date - now, 0), MAX_RETRY_WAIT_MS);
    return fallback;
}

export class GraphError extends Error {
    /** `service` names the upstream in user-facing errors ("SharePoint" for Graph, or "Dataverse"). */
    constructor(public status: number, message: string, public code?: string, public service = "SharePoint") { super(message); }
}

export type RetryDecision = "retry" | "fail";

/**
 * Whether a failed attempt may be repeated.
 *
 * - 429 and 503 mean the service refused the request without processing it: always safe to retry.
 * - Timeouts, network errors, 500/502/504 are ambiguous (the request may have been applied), so they
 *   are retried only for idempotent requests. A path-based upload with conflictBehavior=rename is NOT
 *   idempotent: repeating it after a timeout could create "file 1.pdf" next to "file.pdf".
 */
export function retryDecision(outcome: number | "timeout" | "network", idempotent: boolean, attempt: number): RetryDecision {
    if (attempt >= MAX_ATTEMPTS - 1) return "fail";
    if (outcome === 429 || outcome === 503) return "retry";
    if (outcome === "timeout" || outcome === "network" || outcome === 500 || outcome === 502 || outcome === 504) {
        return idempotent ? "retry" : "fail";
    }
    return "fail";
}

export function isTimeout(e: unknown): boolean {
    const name = (e as Error)?.name;
    return name === "TimeoutError" || name === "AbortError";
}

export const sleep = (ms: number) => new Promise<void>(r => setTimeout(r, ms));

const GRAPH_SCOPE = "https://graph.microsoft.com/.default";

/** App-only token for a resource scope. @azure/identity caches and refreshes tokens per scope. */
async function token(scope: string): Promise<string> {
    credential ??= new ClientSecretCredential(config.tenantId, config.clientId, config.clientSecret);
    const t = await credential.getToken(scope);
    return t.token;
}

export interface GraphInit extends RequestInit {
    json?: unknown;
    /**
     * May the request be repeated after an ambiguous failure? Defaults to true for GET/HEAD/DELETE/PATCH
     * and false for POST/PUT; pass explicitly when a POST/PUT is known to be safe.
     */
    idempotent?: boolean;
    timeoutMs?: number;
}

/**
 * Calls Microsoft Graph with an app-only token. Every attempt has a timeout; throttling and transient
 * failures are retried per retryDecision(). `pathOrUrl` may be relative to /v1.0 or an absolute nextLink.
 * Always throws GraphError on failure (504 for timeouts, 502 for network errors).
 */
export function graph<T = unknown>(pathOrUrl: string, init: GraphInit = {}): Promise<T> {
    const url = pathOrUrl.startsWith("https://") ? pathOrUrl : `${GRAPH}${pathOrUrl}`;
    return request<T>(url, init, () => token(GRAPH_SCOPE), "SharePoint");
}

const dvHeaders = (init: GraphInit) => ({
    Accept: "application/json", "OData-Version": "4.0", "OData-MaxVersion": "4.0",
    ...(init.headers as Record<string, string>),
});
const dvOrg = () => config.dataverseUrl.replace(/\/+$/, "");

/**
 * Calls the Dataverse Web API (v9.2) as the backend's application user, with the same timeout/retry rules.
 * `path` is relative to /api/data/v9.2, e.g. "/ce_documentmanagerconfigs?$filter=...".
 */
export function dataverse<T = unknown>(path: string, init: GraphInit = {}): Promise<T> {
    const org = dvOrg();
    return request<T>(`${org}/api/data/v9.2${path}`, { ...init, headers: dvHeaders(init) }, () => token(`${org}/.default`), "Dataverse");
}

// ── Dataverse as the signed-in user (On-Behalf-Of) ────────────────────────────────────────────────
// The caller's API token is exchanged for a Dataverse token for the same user, so Dataverse applies
// that user's own security roles. Credentials are cached per incoming token (they cache the exchanged
// token internally), bounded so a busy worker can't grow without limit.
const OBO_CACHE_MAX = 500;
const oboCache = new Map<string, OnBehalfOfCredential>();

function oboCredential(userAssertion: string): OnBehalfOfCredential {
    const key = createHash("sha256").update(userAssertion).digest("base64");
    let cred = oboCache.get(key);
    if (!cred) {
        if (oboCache.size >= OBO_CACHE_MAX) oboCache.delete(oboCache.keys().next().value!);
        cred = new OnBehalfOfCredential({
            tenantId: config.tenantId, clientId: config.clientId, clientSecret: config.clientSecret,
            userAssertionToken: userAssertion,
        });
        oboCache.set(key, cred);
    }
    return cred;
}

/** Calls the Dataverse Web API as the user who sent `userAssertion` (their bearer token for this API). */
export function dataverseAsUser<T = unknown>(userAssertion: string, path: string, init: GraphInit = {}): Promise<T> {
    const org = dvOrg();
    const getToken = async () => {
        try {
            return (await oboCredential(userAssertion).getToken(`${org}/.default`)).token;
        } catch (e) {
            // Usually: missing delegated "Dynamics CRM / user_impersonation" consent on the backend app.
            throw new GraphError(502, `Could not get a Dataverse token for the user: ${(e as Error).message}`, "obo", "Dataverse");
        }
    };
    return request<T>(`${org}/api/data/v9.2${path}`, { ...init, headers: dvHeaders(init) }, getToken, "Dataverse");
}

async function request<T>(url: string, init: GraphInit, getToken: () => Promise<string>, service: string): Promise<T> {
    const method = (init.method ?? "GET").toUpperCase();
    const idempotent = init.idempotent ?? ["GET", "HEAD", "DELETE", "PATCH"].includes(method);
    const timeoutMs = init.timeoutMs ?? REQUEST_TIMEOUT_MS;
    const { json, idempotent: _i, timeoutMs: _t, ...fetchInit } = init;

    let body = fetchInit.body;
    const extraHeaders: Record<string, string> = { ...(fetchInit.headers as Record<string, string>) };
    if (json !== undefined) {
        extraHeaders["Content-Type"] = "application/json";
        body = JSON.stringify(json);
    }

    for (let attempt = 0; ; attempt++) {
        let res: Response;
        try {
            // Token fetched per attempt: a long retry sequence must not reuse an expired token.
            const headers = { Authorization: `Bearer ${await getToken()}`, ...extraHeaders };
            res = await fetch(url, { ...fetchInit, method, headers, body, signal: AbortSignal.timeout(timeoutMs) });
        } catch (e) {
            const outcome = isTimeout(e) ? "timeout" : "network";
            if (retryDecision(outcome, idempotent, attempt) === "retry") {
                await sleep(retryDelayMs(null, attempt));
                continue;
            }
            throw outcome === "timeout"
                ? new GraphError(504, `${service} did not respond within ${timeoutMs / 1000}s`, "timeout", service)
                : new GraphError(502, `Could not reach ${service}: ${(e as Error).message}`, "network", service);
        }

        if (!res.ok && retryDecision(res.status, idempotent, attempt) === "retry") {
            await res.body?.cancel().catch(() => undefined);
            await sleep(retryDelayMs(res.headers.get("Retry-After"), attempt));
            continue;
        }
        if (!res.ok) {
            let message = res.statusText, code: string | undefined;
            try {
                const e = (await res.json()) as { error?: { message?: string; code?: string } };
                message = e.error?.message ?? message; code = e.error?.code;
            } catch { /* non-JSON error */ }
            throw new GraphError(res.status, message, code, service);
        }
        if (res.status === 204) return undefined as T;
        const ct = res.headers.get("content-type") ?? "";
        return (ct.includes("application/json") ? await res.json() : undefined) as T;
    }
}
