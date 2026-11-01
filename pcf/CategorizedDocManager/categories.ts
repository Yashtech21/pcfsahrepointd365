/**
 * Category -> SharePoint folder mapping (relative to the contact root folder).
 * The value of `value` is what gets written to the DocCategory choice column.
 */
export interface ICategory {
    value: string;
    label: string;
    folder: string;
}

export const ALL_CATEGORIES: Record<string, ICategory> = {
    "Offer Letter": { value: "Offer Letter", label: "Offer Letter", folder: "Offer Letters" },
    Commissions: { value: "Commissions", label: "Commissions", folder: "Commissions" },
    Payslip: { value: "Payslip", label: "Payslip", folder: "Onboarding Documents/Payslips" },
    "Resignation Letter": { value: "Resignation Letter", label: "Resignation Letter", folder: "Onboarding Documents/Resignation Letters" },
    FinCEN_U4: { value: "FinCEN_U4", label: "FinCEN / U4 Verification", folder: "Onboarding Documents/FinCEN_U4" },
    Misc: { value: "Misc", label: "Miscellaneous", folder: "Miscellaneous" },
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
    return keys.map(k => ALL_CATEGORIES[k] ?? { value: k, label: k, folder: k });
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
