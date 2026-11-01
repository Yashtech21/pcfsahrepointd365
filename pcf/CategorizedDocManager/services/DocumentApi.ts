import { IDocumentItem } from "../types";
import { AuthService } from "./AuthService";

export interface IUploadRequest {
    file: File;
    category: string;
    contactId: string;
    contactName: string;
    sourceEntity: string;
    sourceRecordId: string;
}

/** Thin client for the Azure Function document API (see /backend). */
export class DocumentApi {
    constructor(private baseUrl: string, private auth: AuthService) {
        this.baseUrl = baseUrl.replace(/\/+$/, "");
    }

    private async call<T>(path: string, init: RequestInit = {}): Promise<T> {
        const token = await this.auth.getToken();
        const res = await fetch(`${this.baseUrl}${path}`, {
            ...init,
            headers: { Authorization: `Bearer ${token}`, ...(init.headers ?? {}) },
        });
        if (!res.ok) {
            let msg = `${res.status} ${res.statusText}`;
            try { msg = ((await res.json()) as { error?: string }).error ?? msg; } catch { /* ignore */ }
            throw new Error(msg);
        }
        return (res.status === 204 ? undefined : await res.json()) as T;
    }

    public listByContact(contactId: string): Promise<IDocumentItem[]> {
        return this.call<IDocumentItem[]>(`/documents?contactId=${encodeURIComponent(contactId)}`);
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
        });
    }

    public async previewUrl(itemId: string): Promise<string> {
        return (await this.call<{ url: string }>(`/documents/${encodeURIComponent(itemId)}/preview`, { method: "POST" })).url;
    }

    public async downloadUrl(itemId: string): Promise<string> {
        return (await this.call<{ url: string }>(`/documents/${encodeURIComponent(itemId)}/download`)).url;
    }

    public remove(itemId: string): Promise<void> {
        return this.call<void>(`/documents/${encodeURIComponent(itemId)}`, { method: "DELETE" });
    }
}
