/**
 * Categories the control can offer. `value` is written to the DocCategory choice column.
 * Folder placement is decided server-side (backend/src/lib/categories.ts), which is the source of
 * truth: a new category must be added there first, then here.
 */
export interface ICategory {
    value: string;
    label: string;
}

export const ALL_CATEGORIES: Record<string, ICategory> = {
    "Offer Letter": { value: "Offer Letter", label: "Offer Letter" },
    Commissions: { value: "Commissions", label: "Commissions" },
    Payslip: { value: "Payslip", label: "Payslip" },
    "Resignation Letter": { value: "Resignation Letter", label: "Resignation Letter" },
    FinCEN_U4: { value: "FinCEN_U4", label: "FinCEN / U4 Verification" },
    Misc: { value: "Misc", label: "Miscellaneous" },
};

const DEFAULT_BY_ENTITY: Record<string, string[]> = {
    opportunity: ["Offer Letter", "Commissions", "Misc"],
    "*": ["Payslip", "Resignation Letter", "FinCEN_U4", "Misc"],
};

export function categoriesFor(entityName: string, overrideJson?: string | null): ICategory[] {
    let map = DEFAULT_BY_ENTITY;
    if (overrideJson && overrideJson.trim()) {
        try {
            map = { ...DEFAULT_BY_ENTITY, ...(JSON.parse(overrideJson) as Record<string, string[]>) };
        } catch {
            console.warn("CategorizedDocManager: invalid categoryConfig JSON, using defaults");
        }
    }
    const keys = map[entityName.toLowerCase()] ?? map["*"];
    const unknown = keys.filter(k => !ALL_CATEGORIES[k]);
    if (unknown.length) {
        // The server would reject uploads for these, so don't offer them.
        console.warn(`CategorizedDocManager: ignoring unknown categories in categoryConfig: ${unknown.join(", ")}`);
    }
    const result = keys.filter(k => ALL_CATEGORIES[k]).map(k => ALL_CATEGORIES[k]);
    return result.length ? result : (DEFAULT_BY_ENTITY[entityName.toLowerCase()] ?? DEFAULT_BY_ENTITY["*"]).map(k => ALL_CATEGORIES[k]);
}

export function categoryLabel(value: string): string {
    return ALL_CATEGORIES[value]?.label ?? value;
}

/** Badge colours roughly matching the mockup (background, foreground). */
const COLORS: Record<string, [string, string]> = {
    "Offer Letter": ["#0f6cbd", "#ffffff"],
    Commissions: ["#8a3ab9", "#ffffff"],
    Payslip: ["#107c10", "#ffffff"],
    "Resignation Letter": ["#c50f1f", "#ffffff"],
    FinCEN_U4: ["#ca5010", "#ffffff"],
    Misc: ["#e1dfdd", "#323130"],
};

export function categoryColor(value: string): { backgroundColor: string; color: string } {
    const [backgroundColor, color] = COLORS[value] ?? ["#edebe9", "#323130"];
    return { backgroundColor, color };
}
