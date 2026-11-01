import { GraphError, isTimeout, retryDelayMs, sleep } from "./graph";

/** Chunk size must be a multiple of 320 KiB (Graph requirement). 16 × 320 KiB = 5 MiB. */
export const CHUNK_SIZE = 16 * 320 * 1024;
export const CHUNK_TIMEOUT_MS = 120000;
export const CONTROL_TIMEOUT_MS = 10000;
const MAX_FAILURES = 5;

export interface UploadSessionOptions {
    chunkSize?: number;
    chunkTimeoutMs?: number;
    fetchImpl?: typeof fetch;
    sleepImpl?: (ms: number) => Promise<void>;
}

/**
 * Uploads `content` to a Graph upload session.
 *
 * Reliability rules:
 * - Throttling (429/503) waits for Retry-After and resends the same range.
 * - Ambiguous failures (timeout, network error, 5xx, 416) ask the session for `nextExpectedRanges`
 *   and resume from there, so a chunk that actually landed is never sent twice.
 * - After MAX_FAILURES consecutive failures, or any non-retryable error, the session is cancelled
 *   (with its own timeout) and a GraphError is thrown: 504 for timeouts, otherwise 502/4xx.
 * - The pre-authenticated uploadUrl never receives the bearer token.
 */
export async function uploadToSession<T>(uploadUrl: string, content: Uint8Array, opts: UploadSessionOptions = {}): Promise<T> {
    const doFetch = opts.fetchImpl ?? fetch;
    const wait = opts.sleepImpl ?? sleep;
    const chunkSize = opts.chunkSize ?? CHUNK_SIZE;
    const chunkTimeout = opts.chunkTimeoutMs ?? CHUNK_TIMEOUT_MS;
    const total = content.length;

    const cancel = () => doFetch(uploadUrl, { method: "DELETE", signal: AbortSignal.timeout(CONTROL_TIMEOUT_MS) })
        .then(r => r.body?.cancel(), () => undefined)
        .catch(() => undefined);

    /** Where does the server want us to continue? undefined = session gone. */
    const nextOffset = async (): Promise<number | undefined> => {
        try {
            const r = await doFetch(uploadUrl, { signal: AbortSignal.timeout(CONTROL_TIMEOUT_MS) });
            if (!r.ok) { await r.body?.cancel(); return undefined; }
            const s = (await r.json()) as { nextExpectedRanges?: string[] };
            const first = s.nextExpectedRanges?.[0];
            if (!first) return undefined;
            const start = Number(first.split("-")[0]);
            return Number.isFinite(start) && start >= 0 && start <= total ? start : undefined;
        } catch {
            return undefined;
        }
    };

    let offset = 0;
    let failures = 0;
    let lastError: GraphError = new GraphError(502, "Upload failed");

    try {
        while (offset < total) {
            const end = Math.min(offset + chunkSize, total);
            let res: Response;
            try {
                res = await doFetch(uploadUrl, {
                    method: "PUT",
                    signal: AbortSignal.timeout(chunkTimeout),
                    headers: { "Content-Length": String(end - offset), "Content-Range": `bytes ${offset}-${end - 1}/${total}` },
                    body: content.slice(offset, end),
                });
            } catch (e) {
                lastError = isTimeout(e)
                    ? new GraphError(504, `Upload chunk at byte ${offset} timed out after ${chunkTimeout / 1000}s`, "timeout")
                    : new GraphError(502, `Upload chunk at byte ${offset} failed: ${(e as Error).message}`, "network");
                if (++failures >= MAX_FAILURES) throw lastError;
                await wait(retryDelayMs(null, failures));
                const resume = await nextOffset();
                if (resume === undefined) throw lastError;
                offset = resume;
                continue;
            }

            if (res.status === 200 || res.status === 201) return (await res.json()) as T;   // last chunk: the item
            if (res.status === 202) {                                                        // chunk accepted
                failures = 0;
                const s = (await res.json().catch(() => ({}))) as { nextExpectedRanges?: string[] };
                const next = Number(s.nextExpectedRanges?.[0]?.split("-")[0]);
                offset = Number.isFinite(next) && next > offset ? next : end;
                continue;
            }

            const text = await res.text().catch(() => "");
            lastError = new GraphError(res.status === 504 ? 504 : res.status >= 500 ? 502 : res.status,
                `Upload chunk at byte ${offset} failed (${res.status}): ${text.slice(0, 300)}`);
            if (res.status === 404) throw lastError;                                         // session expired/cancelled
            if (++failures >= MAX_FAILURES) throw lastError;

            if (res.status === 429 || res.status === 503) {
                await wait(retryDelayMs(res.headers.get("Retry-After"), failures));
                continue;                                                                    // same range again
            }
            if (res.status >= 500 || res.status === 416) {
                await wait(retryDelayMs(res.headers.get("Retry-After"), failures));
                const resume = await nextOffset();
                if (resume === undefined) throw lastError;
                offset = resume;
                continue;
            }
            throw lastError;                                                                 // other 4xx: not retryable
        }
        throw new GraphError(502, "Upload session ended without returning the file");
    } catch (e) {
        await cancel();
        throw e instanceof GraphError ? e : new GraphError(502, (e as Error).message);
    }
}
