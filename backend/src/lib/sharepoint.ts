import { config } from "./config";
import { graph, GraphError } from "./graph";
import { contactFolderName, encodePath, isOfficeFile, sanitizeName } from "./naming";

/** Shape returned to the PCF control (mirrors IDocumentItem on the client). */
export interface DocumentDto {
    id: string;
    name: string;
    category: string;
    sourceEntity: string;
    sourceRecordId: string;
    contactId: string;
    uploadDate: string;
    webUrl: string;
    size: number;
    fileType: string;
}

interface Target { siteId: string; listId: string; driveId: string }
interface DriveItem {
    id: string; name: string; webUrl: string; size?: number; createdDateTime: string;
    file?: { mimeType?: string }; folder?: unknown;
    "@microsoft.graph.downloadUrl"?: string;
    listItem?: { fields?: Record<string, unknown> };
}

const SMALL_UPLOAD_LIMIT = 4 * 1024 * 1024;
const CHUNK_SIZE = 16 * 320 * 1024; // 5 MiB, must be a multiple of 320 KiB
const FIELDS = "DocCategory,SourceEntity,SourceRecordId,ContactGuid";

let targetPromise: Promise<Target> | null = null;
const contactFolderCache = new Map<string, string>();

/** Resolves (and caches) the site, the Contacts list and its drive. */
export function getTarget(): Promise<Target> {
    targetPromise ??= (async () => {
        const site = await graph<{ id: string }>(`/sites/${config.spHostname}:${config.spSitePath}`);
        const lists = await graph<{ value: { id: string; name: string; displayName: string }[] }>(
            `/sites/${site.id}/lists?$select=id,name,displayName`);
        const wanted = config.spLibraryName.toLowerCase();
        const list = lists.value.find(l => l.displayName.toLowerCase() === wanted || l.name.toLowerCase() === wanted);
        if (!list) throw new Error(`Document library "${config.spLibraryName}" not found on ${config.spSitePath}`);
        const drive = await graph<{ id: string }>(`/sites/${site.id}/lists/${list.id}/drive?$select=id`);
        return { siteId: site.id, listId: list.id, driveId: drive.id };
    })().catch(e => { targetPromise = null; throw e; });
    return targetPromise;
}

function str(v: unknown): string {
    return typeof v === "string" ? v : "";
}

export function toDto(item: DriveItem, fields: Record<string, unknown> = item.listItem?.fields ?? {}): DocumentDto {
    const ext = item.name.includes(".") ? item.name.split(".").pop()!.toLowerCase() : "";
    return {
        id: item.id,
        name: item.name,
        category: str(fields.DocCategory),
        sourceEntity: str(fields.SourceEntity) || "contact",
        sourceRecordId: str(fields.SourceRecordId).toLowerCase(),
        contactId: str(fields.ContactGuid).toLowerCase(),
        uploadDate: item.createdDateTime,
        webUrl: item.webUrl,
        size: item.size ?? 0,
        fileType: ext,
    };
}

/**
 * Returns the drive-relative path of the contact's root folder, creating it if needed.
 * Folders are matched on the "_{guid}" suffix so a renamed contact keeps using its existing folder.
 */
export async function ensureContactFolder(contactId: string, contactName: string): Promise<string> {
    const cached = contactFolderCache.get(contactId);
    if (cached) return cached;
    const { driveId } = await getTarget();
    const desired = contactFolderName(contactName, contactId);

    const remember = (name: string) => { contactFolderCache.set(contactId, name); return name; };

    try {
        await graph(`/drives/${driveId}/root:/${encodePath(desired)}?$select=id`);
        return remember(desired);
    } catch (e) {
        if (!(e instanceof GraphError && e.status === 404)) throw e;
    }

    // Contact may have been renamed: look for an existing folder ending in _{guid}.
    const found = await graph<{ value: DriveItem[] }>(
        `/drives/${driveId}/root/search(q='${contactId}')?$select=id,name,folder,parentReference&$top=50`);
    const existing = found.value.find(i => i.folder && i.name.toLowerCase().endsWith(`_${contactId}`));
    if (existing) return remember(existing.name);

    try {
        await graph(`/drives/${driveId}/root/children`, {
            method: "POST",
            json: { name: desired, folder: {}, "@microsoft.graph.conflictBehavior": "fail" },
        });
    } catch (e) {
        if (!(e instanceof GraphError && e.status === 409)) throw e; // created concurrently
    }
    return remember(desired);
}

export interface UploadInput {
    fileName: string;
    content: Buffer;
    categoryFolder: string;
    category: string;
    contactId: string;
    contactName: string;
    sourceEntity: string;
    sourceRecordId: string;
}

export async function uploadDocument(input: UploadInput): Promise<DocumentDto> {
    const { driveId } = await getTarget();
    const root = await ensureContactFolder(input.contactId, input.contactName);
    const path = encodePath(`${root}/${input.categoryFolder}/${sanitizeName(input.fileName)}`);

    let item: DriveItem;
    if (input.content.length <= SMALL_UPLOAD_LIMIT) {
        // Graph creates any missing intermediate folders for path-based uploads.
        item = await graph<DriveItem>(`/drives/${driveId}/root:/${path}:/content?@microsoft.graph.conflictBehavior=rename`, {
            method: "PUT",
            headers: { "Content-Type": "application/octet-stream" },
            body: new Uint8Array(input.content),
        });
    } else {
        item = await uploadLarge(driveId, path, input.content);
    }

    const fields = {
        DocCategory: input.category,
        SourceEntity: input.sourceEntity,
        SourceRecordId: input.sourceRecordId,
        ContactGuid: input.contactId,
    };
    await graph(`/drives/${driveId}/items/${item.id}/listItem/fields`, { method: "PATCH", json: fields });
    return toDto(item, fields);
}

async function uploadLarge(driveId: string, path: string, content: Buffer): Promise<DriveItem> {
    const session = await graph<{ uploadUrl: string }>(`/drives/${driveId}/root:/${path}:/createUploadSession`, {
        method: "POST",
        json: { item: { "@microsoft.graph.conflictBehavior": "rename" } },
    });
    const total = content.length;
    for (let start = 0; start < total; start += CHUNK_SIZE) {
        const end = Math.min(start + CHUNK_SIZE, total);
        // uploadUrl is pre-authenticated: it must NOT receive the bearer token.
        const res = await fetch(session.uploadUrl, {
            method: "PUT",
            headers: { "Content-Length": String(end - start), "Content-Range": `bytes ${start}-${end - 1}/${total}` },
            body: new Uint8Array(content.subarray(start, end)),
        });
        if (!res.ok) {
            await fetch(session.uploadUrl, { method: "DELETE" }).catch(() => undefined);
            throw new GraphError(res.status, `Chunk upload failed at byte ${start}: ${await res.text()}`);
        }
        if (end === total) return (await res.json()) as DriveItem;
    }
    throw new Error("Upload session ended unexpectedly");
}

/** All documents tagged with the contact GUID (indexed column), across every source entity. */
export async function listByContact(contactId: string): Promise<DocumentDto[]> {
    const { siteId, listId } = await getTarget();
    const results: DocumentDto[] = [];
    let url: string | undefined =
        `/sites/${siteId}/lists/${listId}/items` +
        `?$expand=fields($select=${FIELDS}),driveItem($select=id,name,webUrl,size,createdDateTime,file)` +
        `&$filter=fields/ContactGuid eq '${contactId}'&$top=200`;

    while (url) {
        const page: { value: { fields?: Record<string, unknown>; driveItem?: DriveItem }[]; "@odata.nextLink"?: string } =
            await graph(url, { headers: { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" } });
        for (const li of page.value) {
            if (li.driveItem?.file) results.push(toDto(li.driveItem, li.fields));
        }
        url = page["@odata.nextLink"];
    }
    return results;
}

export async function getItem(itemId: string): Promise<DriveItem> {
    const { driveId } = await getTarget();
    return graph<DriveItem>(`/drives/${driveId}/items/${encodeURIComponent(itemId)}?$expand=listItem($expand=fields($select=${FIELDS}))`);
}

/** Short-lived embeddable preview URL (Graph driveItem: preview). */
export async function previewUrl(item: DriveItem): Promise<string> {
    const { driveId } = await getTarget();
    const body = isOfficeFile(item.name) ? { viewer: "office", chromeless: true } : {};
    const r = await graph<{ getUrl?: string }>(`/drives/${driveId}/items/${item.id}/preview`, { method: "POST", json: body });
    if (!r.getUrl) throw new Error("Preview is not available for this file type");
    return r.getUrl;
}

export async function downloadUrl(itemId: string): Promise<string> {
    const { driveId } = await getTarget();
    const item = await graph<DriveItem>(`/drives/${driveId}/items/${encodeURIComponent(itemId)}`);
    const url = item["@microsoft.graph.downloadUrl"];
    if (!url) throw new Error("Item is not a downloadable file");
    return url;
}

export async function deleteItem(itemId: string): Promise<void> {
    const { driveId } = await getTarget();
    await graph(`/drives/${driveId}/items/${encodeURIComponent(itemId)}`, { method: "DELETE" });
}
