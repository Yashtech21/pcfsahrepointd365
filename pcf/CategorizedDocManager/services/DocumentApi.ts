import { IDocumentItem } from "../types";
import { AuthService } from "./AuthService";
import { fetchWithTimeout } from "./http";

export interface IUploadRequest {
    file: File;
    category: string;
    contactId: string;
    contactName: string;
    sourceEntity: string;
    sourceRecordId: string;
}

const DEFAULT_TIMEOUT_MS = 8000;
const PREVIEW_TIMEOUT_MS = 15000;
/** Uploads get 60s plus time for the payload at a conservative ~250 KB/s. */
const uploadTimeoutMs = (bytes: number) => 60000 + Math.ceil(bytes / 256);

/** Thin client for the Azure Function document API (see /backend). */
export class DocumentApi {
    private baseUrl: string;
    /** In-flight list requests, so repeated updateView/effect runs share one Graph round trip. */
    private pendingLists = new Map<string, Promise<IDocumentItem[]>>();

    constructor(baseUrl: string, private auth: AuthService) {
        this.baseUrl = baseUrl.replace(/\/+$/, "");
    }

    private async call<T>(path: string, init: RequestInit = {}, timeoutMs = DEFAULT_TIMEOUT_MS): Promise<T> {
        const token = await this.auth.getToken();
        const res = await fetchWithTimeout(`${this.baseUrl}${path}`, {
            ...init,
            headers: { Authorization: `Bearer ${token}`, ...(init.headers ?? {}) },
        }, timeoutMs);
        if (!res.ok) {
            let msg = `${res.status} ${res.statusText}`;
            try { msg = ((await res.json()) as { error?: string }).error ?? msg; } catch { /* ignore */ }
            throw new Error(msg);
        }
        return (res.status === 204 ? undefined : await res.json()) as T;
    }

    public listByContact(contactId: string): Promise<IDocumentItem[]> {
        const pending = this.pendingLists.get(contactId);
        if (pending) return pending;
        const p = this.call<IDocumentItem[]>(`/documents?contactId=${encodeURIComponent(contactId)}`)
            .finally(() => this.pendingLists.delete(contactId));
        this.pendingLists.set(contactId, p);
        return p;
    }

    public upload(req: IUploadRequest): Promise<IDocumentItem> {
        const q = new URLSearchParams({
            fileName: req.file.name,
            category: req.category,
            contactId: req.contactId,
            contactName: req.contactName,
            sourceEntity: req.sourceEntity,
            sourceRecordId: req.sourceRecordId,
        });
        return this.call<IDocumentItem>(`/documents/upload?${q.toString()}`, {
            method: "POST",
            headers: { "Content-Type": req.file.type || "application/octet-stream" },
            body: req.file,
        }, uploadTimeoutMs(req.file.size));
    }

    public async previewUrl(itemId: string): Promise<string> {
        return (await this.call<{ url: string }>(`/documents/${encodeURIComponent(itemId)}/preview`, { method: "POST" }, PREVIEW_TIMEOUT_MS)).url;
    }

    public async downloadUrl(itemId: string): Promise<string> {
        return (await this.call<{ url: string }>(`/documents/${encodeURIComponent(itemId)}/download`)).url;
    }

    /** `confirmedFileName` must be the name the user confirmed; the server refuses mismatches. */
    public remove(itemId: string, confirmedFileName: string): Promise<void> {
        return this.call<void>(`/documents/${encodeURIComponent(itemId)}`, {
            method: "DELETE",
            headers: { "X-Confirm-Delete": encodeURIComponent(confirmedFileName) },
        });
    }
}
