import * as React from "react";

/**
 * Categories come from Dataverse (ce_documentcategory), loaded by SettingsService. Nothing here is
 * hard-coded any more: admins add, rename, recolour or deactivate categories in the app.
 * `value` is what is stored in SharePoint (DocCategory); the backend decides the folder.
 */
export interface ICategory {
    value: string;
    label: string;
    color?: string;
    sortOrder: number;
}

export interface ICategoryCatalog {
    label(value: string): string;
    color(value: string): { backgroundColor: string; color: string };
}

/** Used when a category has no colour configured: stable per value, all readable with white text. */
const FALLBACK_PALETTE = ["#0f6cbd", "#8a3ab9", "#107c10", "#c50f1f", "#ca5010", "#038387", "#5c2e91", "#986f0b"];

function hash(s: string): number {
    let h = 0;
    for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) | 0;
    return Math.abs(h);
}

/** Picks black or white text for a hex background (WCAG relative luminance). */
function textOn(hex: string): string {
    const m = /^#([0-9a-f]{6})$/i.exec(hex);
    if (!m) return "#ffffff";
    const n = parseInt(m[1], 16);
    const lin = (c: number) => { const v = c / 255; return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4; };
    const l = 0.2126 * lin((n >> 16) & 255) + 0.7152 * lin((n >> 8) & 255) + 0.0722 * lin(n & 255);
    return l > 0.4 ? "#242424" : "#ffffff";
}

export function createCatalog(categories: ICategory[]): ICategoryCatalog {
    const byValue = new Map(categories.map(c => [c.value, c]));
    return {
        // Unknown values (e.g. a deleted category, or a file tagged in SharePoint) show their raw value.
        label: (v) => byValue.get(v)?.label ?? v,
        color: (v) => {
            if (!v) return { backgroundColor: "#e1dfdd", color: "#323130" };
            const bg = byValue.get(v)?.color || FALLBACK_PALETTE[hash(v) % FALLBACK_PALETTE.length];
            return { backgroundColor: bg, color: textOn(bg) };
        },
    };
}

export const CategoryCatalogContext = React.createContext<ICategoryCatalog>(createCatalog([]));

export const useCategoryCatalog = () => React.useContext(CategoryCatalogContext);
