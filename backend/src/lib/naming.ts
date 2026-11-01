const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;

export function normalizeGuid(value: string | null | undefined): string | null {
    const g = (value ?? "").replace(/[{}]/g, "").trim().toLowerCase();
    return GUID_RE.test(g) ? g : null;
}

/** Removes characters SharePoint does not allow in file/folder names and trims trailing dots/spaces. */
export function sanitizeName(name: string): string {
    const cleaned = name
        .replace(/["*:<>?/\\|#%]/g, "_")
        .replace(/[\u0000-\u001f]/g, "")
        .replace(/\s+/g, " ")
        .trim()
        .replace(/[. ]+$/, "");
    return cleaned.slice(0, 200) || "_";
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
