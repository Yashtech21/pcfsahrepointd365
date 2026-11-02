import { config } from "./config";
import { DeleteNotConfirmedError, UnsupportedItemError } from "./errors";
import { assertFileAllowed } from "./fileRules";
import { graph, GraphError } from "./graph";
import { contactFolderName, encodePath, isOfficeFile, sanitizeName } from "./naming";
import { uploadToSession } from "./uploadSession";

export { DeleteNotConfirmedError, UnsupportedItemError };

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
const FOLDER_CACHE_MAX = 5000;
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

/** File links go to the browser ("Open in SharePoint", "Copy link"): only trusted https hosts, else none. */
function safeWebUrl(raw: string | undefined): string {
    if (!raw) return "";
    try { return assertTrustedUrl(raw); } catch { return ""; }
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
        webUrl: safeWebUrl(item.webUrl),
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

    const remember = (name: string) => {
        // Bounded: a long-lived worker serving many contacts must not grow without limit.
        if (contactFolderCache.size >= FOLDER_CACHE_MAX) contactFolderCache.delete(contactFolderCache.keys().next().value!);
        contactFolderCache.set(contactId, name);
        return name;
    };

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
            idempotent: true, // conflictBehavior=fail: a repeat yields 409, handled below
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
    const fileName = sanitizeName(input.fileName);
    // Validate before touching SharePoint: blocked types, executables, content/extension mismatch.
    assertFileAllowed(fileName, input.content);

    const { driveId } = await getTarget();
    const root = await ensureContactFolder(input.contactId, input.contactName);
    const path = encodePath(`${root}/${input.categoryFolder}/${fileName}`);
    const bytes = new Uint8Array(input.content.buffer, input.content.byteOffset, input.content.length);

    let item: DriveItem;
    if (bytes.length <= SMALL_UPLOAD_LIMIT) {
        // Graph creates any missing intermediate folders for path-based uploads. Not idempotent
        // (conflictBehavior=rename), so graph() won't repeat it after a timeout.
        item = await graph<DriveItem>(`/drives/${driveId}/root:/${path}:/content?@microsoft.graph.conflictBehavior=rename`, {
            method: "PUT",
            headers: { "Content-Type": "application/octet-stream" },
            body: bytes as unknown as BodyInit, // type-only: Node accepts any Uint8Array
        });
    } else {
        const session = await graph<{ uploadUrl: string }>(`/drives/${driveId}/root:/${path}:/createUploadSession`, {
            method: "POST",
            idempotent: true, // an unused duplicate session simply expires
            json: { item: { "@microsoft.graph.conflictBehavior": "rename" } },
        });
        item = await uploadToSession<DriveItem>(session.uploadUrl, bytes);
    }

    const rollback = () => graph(`/drives/${driveId}/items/${item.id}`, { method: "DELETE" }).catch(() => undefined);

    // Integrity: SharePoint must hold exactly what we received.
    if (typeof item.size === "number" && item.size !== bytes.length) {
        await rollback();
        throw new GraphError(502, `Stored file size (${item.size}) does not match the upload (${bytes.length}); the upload was discarded`);
    }

    const fields = {
        DocCategory: input.category,
        SourceEntity: input.sourceEntity,
        SourceRecordId: input.sourceRecordId,
        ContactGuid: input.contactId,
    };
    try {
        await graph(`/drives/${driveId}/items/${item.id}/listItem/fields`, { method: "PATCH", json: fields });
    } catch (e) {
        // Without ContactGuid the file would be invisible to the control and a retry would create a
        // renamed duplicate, so roll the upload back before reporting the failure.
        await rollback();
        throw e;
    }
    return toDto(item, fields);
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

export interface PreviewTarget {
    url: string;
    /** Present when Graph returns a POST-based embed: form-encoded parameters to POST to `url`. */
    postParameters?: string;
}

/** Only https URLs on Microsoft/SharePoint hosts are handed to the browser for embedding/downloading. */
export function assertTrustedUrl(raw: string): string {
    let u: URL;
    try { u = new URL(raw); } catch { throw new GraphError(502, "Graph returned an invalid URL"); }
    const host = u.hostname.toLowerCase();
    const trusted = [".sharepoint.com", ".sharepoint.us", ".sharepoint.de", ".sharepoint.cn", ".microsoft.com",
        ".office.com", ".office.net", ".officeapps.live.com", ".svc.ms", ".live.com", ".microsoftonline.com"];
    if (u.protocol !== "https:" || !trusted.some(t => host === t.slice(1) || host.endsWith(t))) {
        throw new GraphError(502, `Graph returned an untrusted URL host (${host})`);
    }
    return u.toString();
}

/** A file plus the tags needed to authorize access to it. */
export interface FileMeta {
    id: string;
    name: string;
    isFile: boolean;
    downloadUrl?: string;
    contactId: string;       // ContactGuid ("" if untagged)
    sourceEntity: string;    // SourceEntity, "contact" when empty
    sourceRecordId: string;  // SourceRecordId ("" if none)
}

/** One Graph call: the item, its download URL and its metadata columns. */
export async function getFileMeta(itemId: string): Promise<FileMeta> {
    const { driveId } = await getTarget();
    const item = await graph<DriveItem>(`/drives/${driveId}/items/${encodeURIComponent(itemId)}?$expand=listItem($expand=fields($select=${FIELDS}))`);
    const dto = toDto(item);
    return {
        id: item.id, name: item.name, isFile: !!item.file && !item.folder,
        downloadUrl: item["@microsoft.graph.downloadUrl"],
        contactId: dto.contactId, sourceEntity: dto.sourceEntity, sourceRecordId: dto.sourceRecordId,
    };
}

/** Short-lived embeddable preview (Graph driveItem: preview). */
export async function previewTarget(file: FileMeta): Promise<PreviewTarget> {
    if (!file.isFile) throw new UnsupportedItemError("Only files can be previewed");
    const { driveId } = await getTarget();
    const body = isOfficeFile(file.name) ? { viewer: "office", chromeless: true } : {};
    const r = await graph<{ getUrl?: string; postUrl?: string; postParameters?: string }>(
        `/drives/${driveId}/items/${file.id}/preview`, { method: "POST", idempotent: true, json: body });
    if (r.getUrl) return { url: assertTrustedUrl(r.getUrl) };
    if (r.postUrl) return { url: assertTrustedUrl(r.postUrl), postParameters: r.postParameters ?? "" };
    throw new UnsupportedItemError("Preview is not available for this file type");
}

export function downloadUrl(file: FileMeta): string {
    if (!file.isFile || !file.downloadUrl) throw new UnsupportedItemError("Item is not a downloadable file");
    return assertTrustedUrl(file.downloadUrl);
}

/**
 * Deletes a single file (it goes to the SharePoint recycle bin). Folders are never deleted through
 * the API, and the caller must echo the file name it confirmed, so a stale or forged id cannot
 * remove a different item.
 */
export async function deleteFile(file: FileMeta, confirmedName: string): Promise<void> {
    if (!file.isFile) throw new UnsupportedItemError("Only files can be deleted");
    if (!confirmedName || confirmedName !== file.name) {
        throw new DeleteNotConfirmedError("Delete was not confirmed for this file");
    }
    const { driveId } = await getTarget();
    await graph(`/drives/${driveId}/items/${file.id}`, { method: "DELETE" });
}
