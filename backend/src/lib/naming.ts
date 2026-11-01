const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;

export function normalizeGuid(value: string | null | undefined): string | null {
    const g = (value ?? "").replace(/[{}]/g, "").trim().toLowerCase();
    return GUID_RE.test(g) ? g : null;
}

const RESERVED = /^(con|prn|aux|nul|com[0-9]|lpt[0-9]|desktop\.ini|\.lock|thumbs\.db)(\.|$)/i;

/**
 * Makes a name safe for SharePoint/OneDrive: removes illegal and control characters, the "~$" prefix,
 * "_vti_" (both rejected by SharePoint), leading/trailing dots and spaces, and prefixes reserved device
 * names. The extension is preserved when the name has to be shortened.
 */
export function sanitizeName(name: string): string {
    let cleaned = name
        .normalize("NFC")
        .replace(/["*:<>?/\\|#%]/g, "_")
        .replace(/[\u0000-\u001f\u007f]/g, "")
        .replace(/_vti_/gi, "_vti-")
        .replace(/\s+/g, " ")
        .trim()
        .replace(/^~\$/, "")
        .replace(/^[. ]+/, "")
        .replace(/[. ]+$/, "");
    if (RESERVED.test(cleaned)) cleaned = `_${cleaned}`;
    if (cleaned.length > 200) {
        const dot = cleaned.lastIndexOf(".");
        const ext = dot > 0 && cleaned.length - dot <= 16 ? cleaned.slice(dot) : "";
        cleaned = cleaned.slice(0, 200 - ext.length).replace(/[. ]+$/, "") + ext;
    }
    return cleaned || "_";
}

/** Contact root folder per the spec: {Contact_FullName}_{Contact_GUID} */
export function contactFolderName(contactName: string, contactId: string): string {
    const base = sanitizeName(contactName || "Contact").slice(0, 120);
    return `${base}_${contactId}`;
}

/** Encodes a drive-relative path for the Graph `root:/{path}:` syntax. */
export function encodePath(path: string): string {
    return path.split("/").filter(Boolean).map(encodeURIComponent).join("/");
}

export function isValidEntityName(name: string): boolean {
    return /^[a-z][a-z0-9_]{0,63}$/.test(name);
}

const OFFICE_EXT = new Set(["doc", "docx", "xls", "xlsx", "ppt", "pptx", "odt", "ods", "odp", "rtf"]);
export function isOfficeFile(name: string): boolean {
    const i = name.lastIndexOf(".");
    return i >= 0 && OFFICE_EXT.has(name.slice(i + 1).toLowerCase());
}
