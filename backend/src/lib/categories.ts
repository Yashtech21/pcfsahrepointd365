/**
 * Server-side source of truth for categories. Keys are the values stored in the
 * DocCategory choice column; values are folder paths under the contact root folder.
 * Keep in sync with pcf/CategorizedDocManager/categories.ts.
 */
export const CATEGORY_FOLDERS: Record<string, string> = {
    "Offer Letter": "Offer Letters",
    Commissions: "Commissions",
    Payslip: "Onboarding Documents/Payslips",
    "Resignation Letter": "Onboarding Documents/Resignation Letters",
    FinCEN_U4: "Onboarding Documents/FinCEN_U4",
    Misc: "Miscellaneous",
};

export function folderForCategory(category: string): string | undefined {
    return CATEGORY_FOLDERS[category];
}
