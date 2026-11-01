export class TimeoutError extends Error {
    constructor(seconds: number) {
        super(`The document service did not respond within ${seconds} seconds.`);
        this.name = "TimeoutError";
    }
}

/** fetch() with a hard timeout so a stalled call never leaves the D365 form spinning. */
export async function fetchWithTimeout(url: string, init: RequestInit, timeoutMs: number): Promise<Response> {
    const controller = new AbortController();
    const timer = window.setTimeout(() => controller.abort(), timeoutMs);
    try {
        return await fetch(url, { ...init, signal: controller.signal });
    } catch (e) {
        if (controller.signal.aborted) throw new TimeoutError(Math.round(timeoutMs / 1000));
        throw e;
    } finally {
        window.clearTimeout(timer);
    }
}

/** Rejects with TimeoutError if `promise` does not settle in time (used for token acquisition). */
export function withTimeout<T>(promise: Promise<T>, timeoutMs: number): Promise<T> {
    return new Promise<T>((resolve, reject) => {
        const timer = window.setTimeout(() => reject(new TimeoutError(Math.round(timeoutMs / 1000))), timeoutMs);
        promise.then(
            v => { window.clearTimeout(timer); resolve(v); },
            e => { window.clearTimeout(timer); reject(e); },
        );
    });
}

export interface IXhrResult { status: number; statusText: string; body: string }

/**
 * Upload with byte-level progress. fetch() can't report upload progress, so this uses XMLHttpRequest,
 * with the same timeout contract as fetchWithTimeout (rejects with TimeoutError).
 */
export function xhrUpload(
    url: string,
    opts: { headers: Record<string, string>; body: Blob; timeoutMs: number; onProgress?: (loaded: number, total: number) => void },
): Promise<IXhrResult> {
    return new Promise((resolve, reject) => {
        const xhr = new XMLHttpRequest();
        xhr.open("POST", url);
        Object.entries(opts.headers).forEach(([k, v]) => xhr.setRequestHeader(k, v));
        xhr.timeout = opts.timeoutMs;
        let sent = 0;
        xhr.upload.onprogress = e => {
            sent = e.loaded;
            opts.onProgress?.(e.loaded, e.lengthComputable ? e.total : opts.body.size);
        };
        xhr.onload = () => resolve({ status: xhr.status, statusText: xhr.statusText, body: xhr.responseText });
        xhr.ontimeout = () => reject(new TimeoutError(Math.round(opts.timeoutMs / 1000)));
        xhr.onerror = () => reject(new Error(sent > 0 && sent < opts.body.size
            // The server closed the connection mid-send: typically it rejected the request early
            // (e.g. file too large) before reading the whole body.
            ? `The connection was closed after ${Math.round((sent / opts.body.size) * 100)}% of the file was sent. `
              + "The file may be larger than the server allows, or the network dropped."
            : "Network error while uploading"));
        xhr.onabort = () => reject(new Error("Upload was cancelled"));
        xhr.send(opts.body);
    });
}
