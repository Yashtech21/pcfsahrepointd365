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
