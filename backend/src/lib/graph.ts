import { ClientSecretCredential } from "@azure/identity";
import { config } from "./config";

const GRAPH = "https://graph.microsoft.com/v1.0";
let credential: ClientSecretCredential | null = null;

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
        const res = await fetch(url, { ...init, headers, body });
        if ((res.status === 429 || res.status === 503 || res.status === 504) && attempt < 4) {
            const wait = Number(res.headers.get("Retry-After") ?? 2 ** attempt) * 1000;
            await new Promise(r => setTimeout(r, Math.min(wait, 30000)));
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
