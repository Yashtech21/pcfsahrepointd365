export function normalizeGuid(id?: string | null): string {
    return (id ?? "").replace(/[{}]/g, "").toLowerCase();
}

/**
 * Fallback label used only until (or if) the entity's display name loads from metadata.
 * It does not guess at publisher prefixes: "advisor_onboarding" -> "Advisor Onboarding".
 */
export function prettifyEntityName(logicalName: string): string {
    return logicalName
        .split("_")
        .filter(Boolean)
        .map(w => w.charAt(0).toUpperCase() + w.slice(1))
        .join(" ");
}

export function fileExtension(name: string): string {
    const i = name.lastIndexOf(".");
    return i >= 0 ? name.slice(i + 1).toLowerCase() : "";
}

export function formatSize(bytes: number): string {
    if (!bytes) return "";
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
    return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

export function formatDate(iso: string): string {
    const d = new Date(iso);
    return isNaN(d.getTime()) ? "" : d.toLocaleDateString(undefined, { year: "numeric", month: "2-digit", day: "2-digit" });
}
