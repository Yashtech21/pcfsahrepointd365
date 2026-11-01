import { ClientSecretCredential } from "@azure/identity";
import { config } from "./config";

const GRAPH = "https://graph.microsoft.com/v1.0";
let credential: ClientSecretCredential | null = null;

const REQUEST_TIMEOUT_MS = Number(process.env.GRAPH_TIMEOUT_MS || "30000");
const MAX_RETRY_WAIT_MS = 30000;

/**
 * Retry-After may be delta-seconds or an HTTP-date. Anything unparsable falls back to
 * exponential backoff so a throttled call never retries instantly.
 */
export function retryDelayMs(header: string | null, attempt: number, now = Date.now()): number {
    const fallback = 2 ** attempt * 1000;
    if (!header) return fallback;
    const seconds = Number(header);
    if (Number.isFinite(seconds) && seconds >= 0) return Math.min(seconds * 1000, MAX_RETRY_WAIT_MS);
    const date = Date.parse(header);
    if (Number.isFinite(date)) return Math.min(Math.max(date - now, 0), MAX_RETRY_WAIT_MS);
    return fallback;
}

export class GraphError extends Error {
    constructor(public status: number, message: string, public code?: string) { super(message); }
}

async function token(): Promise<string> {
    credential ??= new ClientSecretCredential(config.tenantId, config.clientId, config.clientSecret);
    const t = await credential.getToken("https://graph.microsoft.com/.default");
    return t.token;
}

/**
 * Calls Microsoft Graph with an app-only token. Retries throttling (429) and transient 5xx
 * responses honouring Retry-After. `pathOrUrl` may be relative to /v1.0 or an absolute nextLink.
 */
export async function graph<T = unknown>(pathOrUrl: string, init: RequestInit & { json?: unknown } = {}): Promise<T> {
    const url = pathOrUrl.startsWith("https://") ? pathOrUrl : `${GRAPH}${pathOrUrl}`;
    const headers: Record<string, string> = { Authorization: `Bearer ${await token()}`, ...(init.headers as Record<string, string>) };
    let body = init.body;
    if (init.json !== undefined) {
        headers["Content-Type"] = "application/json";
        body = JSON.stringify(init.json);
    }

    for (let attempt = 0; ; attempt++) {
        let res: Response;
        try {
            res = await fetch(url, { ...init, headers, body, signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS) });
        } catch (e) {
            if ((e as Error).name === "TimeoutError") {
                if (attempt < 2) continue;
                throw new GraphError(504, `Microsoft Graph did not respond within ${REQUEST_TIMEOUT_MS / 1000}s`, "timeout");
            }
            throw e;
        }
        if ((res.status === 429 || res.status === 503 || res.status === 504) && attempt < 4) {
            await new Promise(r => setTimeout(r, retryDelayMs(res.headers.get("Retry-After"), attempt)));
            continue;
        }
        if (!res.ok) {
            let message = res.statusText, code: string | undefined;
            try {
                const e = (await res.json()) as { error?: { message?: string; code?: string } };
                message = e.error?.message ?? message; code = e.error?.code;
            } catch { /* non-JSON error */ }
            throw new GraphError(res.status, message, code);
        }
        if (res.status === 204) return undefined as T;
        const ct = res.headers.get("content-type") ?? "";
        return (ct.includes("application/json") ? await res.json() : undefined) as T;
    }
}
