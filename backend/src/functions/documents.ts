import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { AccessChecker, AccessDeniedError, accessFor, AccessLevel, RecordContactMismatchError } from "../lib/access";
import { AuthError, Caller, getCaller } from "../lib/auth";
import { config } from "../lib/config";
import { RejectedFileError } from "../lib/errors";
import { GraphError } from "../lib/graph";
import { isValidEntityName, normalizeGuid } from "../lib/naming";
import { ConfigError, planUpload } from "../lib/tableConfig";
import * as sp from "../lib/sharepoint";

type Handler = (req: HttpRequest, ctx: InvocationContext) => Promise<HttpResponseInit>;
type AuthedHandler = (req: HttpRequest, ctx: InvocationContext, caller: Caller, access: AccessChecker) => Promise<HttpResponseInit>;

class BadRequest extends Error {}

const json = (status: number, body: unknown): HttpResponseInit => ({ status, jsonBody: body });

/** Shared auth + error handling for every endpoint. */
function handler(fn: AuthedHandler): Handler {
    return async (req, ctx) => {
        let who = "anonymous";
        try {
            const caller = await getCaller(req);
            who = caller.name || caller.userId;
            ctx.log(`${req.method} ${req.url} by ${who}`);
            return await fn(req, ctx, caller, accessFor(caller));
        } catch (e) {
            if (e instanceof AuthError) return json(e.status, { error: e.message });
            if (e instanceof AccessDeniedError) {
                ctx.warn(`Access denied for ${who} on ${req.method} ${req.url}: ${e.message}`);
                return json(403, { error: e.message });
            }
            if (e instanceof RecordContactMismatchError) return json(409, { error: e.message });
            if (e instanceof BadRequest || e instanceof sp.UnsupportedItemError || e instanceof ConfigError) {
                return json(400, { error: e.message });
            }
            if (e instanceof RejectedFileError) return json(415, { error: e.message });
            if (e instanceof sp.DeleteNotConfirmedError) return json(409, { error: e.message });
            if (e instanceof GraphError) {
                ctx.warn(`Graph error ${e.status} ${e.code ?? ""}: ${e.message}`);
                const status = e.status === 404 ? 404 : e.status === 403 ? 403 : e.status === 504 ? 504
                    : e.status === 413 || e.status === 507 ? 507 : 502;
                return json(status, { error: e.status === 404 ? "Document not found" : `${e.service} error: ${e.message}` });
            }
            ctx.error(e);
            return json(500, { error: "Unexpected server error" });
        }
    };
}

function requireGuid(value: string | null, name: string): string {
    const g = normalizeGuid(value);
    if (!g) throw new BadRequest(`Query parameter ${name} must be a GUID`);
    return g;
}

/**
 * Authorizes an operation on an existing file from its SharePoint tags: the user must read its contact,
 * and have `level` access on its source record (or on the contact for documents without one).
 * Untagged files can't be authorized and are refused.
 */
async function authorizeFile(access: AccessChecker, file: sp.FileMeta, level: AccessLevel): Promise<void> {
    if (!file.contactId) throw new AccessDeniedError("This document isn't linked to a contact, so access to it can't be verified");
    await access.assertContact(file.contactId, "read");
    if (file.sourceRecordId && file.sourceEntity !== "contact") {
        await access.assertRecord(file.sourceEntity, file.sourceRecordId, level);
    } else if (level === "write") {
        await access.assertContact(file.contactId, "write");
    }
}

function requireItemId(req: HttpRequest): string {
    const id = req.params.itemId ?? "";
    if (!/^[A-Za-z0-9!_\-]+$/.test(id)) throw new BadRequest("Invalid item id");
    return id;
}

// GET /api/documents?contactId={guid}
app.http("listDocuments", {
    methods: ["GET"],
    authLevel: "anonymous", // protected by App Service Authentication, see README
    route: "documents",
    handler: handler(async (req, _ctx, _caller, access) => {
        const contactId = requireGuid(req.query.get("contactId"), "contactId");
        await access.assertContact(contactId, "read");
        // Hide documents from records this user can't see (e.g. another team's opportunity).
        return json(200, await access.filterReadable(await sp.listByContact(contactId)));
    }),
});

// POST /api/documents/upload?fileName=&category=&contactId=&contactName=&sourceEntity=&sourceRecordId=
// Body: raw file bytes.
app.http("uploadDocument", {
    methods: ["POST"],
    authLevel: "anonymous",
    route: "documents/upload",
    handler: handler(async (req, _ctx, _caller, access) => {
        const q = req.query;
        const contactId = requireGuid(q.get("contactId"), "contactId");
        const sourceRecordId = requireGuid(q.get("sourceRecordId"), "sourceRecordId");
        const sourceEntity = (q.get("sourceEntity") ?? "").toLowerCase();
        if (!isValidEntityName(sourceEntity)) throw new BadRequest("Invalid sourceEntity");
        const category = q.get("category") ?? "";
        const fileName = (q.get("fileName") ?? "").trim();
        if (!fileName) throw new BadRequest("fileName is required");

        // Table must be configured, category active and allowed for it (Dataverse configuration).
        const plan = await planUpload(sourceEntity, category, config.maxUploadBytes);

        // Per-record security, as the user: they must see the contact, the record must belong to that
        // contact (configured Contact Lookup Column), and they need CHANGE_ACCESS on the record.
        await access.assertContact(contactId, "read");
        await access.assertRecordLinkedToContact(plan.table, sourceRecordId, contactId);
        if (config.changeAccess === "write") await access.assertRecord(sourceEntity, sourceRecordId, "write");
        const tooLarge = () => json(413, { error: `File is too large (limit ${Math.round(plan.maxBytes / 1024 / 1024)} MB for ${plan.table.name})` });

        const declared = Number(req.headers.get("content-length") ?? "0");
        if (declared > plan.maxBytes) return tooLarge();
        const content = Buffer.from(await req.arrayBuffer());
        if (!content.length) throw new BadRequest("Empty file");
        if (content.length > plan.maxBytes) return tooLarge();

        const doc = await sp.uploadDocument({
            fileName, content, category, categoryFolder: plan.folderPath, contactId,
            contactName: q.get("contactName") ?? "", sourceEntity, sourceRecordId,
        });
        return json(201, doc);
    }),
});

// POST /api/documents/{itemId}/preview -> { url, postParameters? }
app.http("previewDocument", {
    methods: ["POST"],
    authLevel: "anonymous",
    route: "documents/{itemId}/preview",
    handler: handler(async (req, _ctx, _caller, access) => {
        const file = await sp.getFileMeta(requireItemId(req));
        await authorizeFile(access, file, "read");
        return json(200, await sp.previewTarget(file));
    }),
});

// GET /api/documents/{itemId}/download -> { url } (short-lived pre-authenticated URL)
app.http("downloadDocument", {
    methods: ["GET"],
    authLevel: "anonymous",
    route: "documents/{itemId}/download",
    handler: handler(async (req, _ctx, _caller, access) => {
        const file = await sp.getFileMeta(requireItemId(req));
        await authorizeFile(access, file, "read");
        return json(200, { url: sp.downloadUrl(file) });
    }),
});

// DELETE /api/documents/{itemId}
// Header X-Confirm-Delete: <url-encoded file name the user confirmed>. Files only; folders are refused.
app.http("deleteDocument", {
    methods: ["DELETE"],
    authLevel: "anonymous",
    route: "documents/{itemId}",
    handler: handler(async (req, ctx, caller, access) => {
        const itemId = requireItemId(req);
        let confirmed = "";
        try { confirmed = decodeURIComponent(req.headers.get("x-confirm-delete") ?? ""); } catch { /* malformed */ }
        if (!confirmed) throw new BadRequest("Missing X-Confirm-Delete header");
        const file = await sp.getFileMeta(itemId);
        await authorizeFile(access, file, config.changeAccess);
        await sp.deleteFile(file, confirmed);
        ctx.log(`Deleted document ${itemId} (${confirmed}) by ${caller.name || caller.userId}`);
        return { status: 204 };
    }),
});
