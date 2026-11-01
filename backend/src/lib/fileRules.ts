import { RejectedFileError } from "./errors";

/**
 * Executable/script types that must never be stored next to HR documents. Override with the
 * BLOCKED_EXTENSIONS app setting (comma-separated, no dots).
 */
const DEFAULT_BLOCKED = [
    "exe", "dll", "com", "scr", "msi", "msp", "bat", "cmd", "ps1", "psm1", "psd1", "vbs", "vbe", "js", "jse",
    "wsf", "wsh", "hta", "cpl", "jar", "lnk", "reg", "sh", "app", "dmg", "iso", "img", "vhd", "vhdx", "url",
    "scf", "inf", "sys", "drv", "ocx", "pif", "gadget", "application", "appx", "msix", "xll", "xlam", "docm",
    "dotm", "xlsm", "xltm", "pptm", "potm", "ppam", "sldm",
];

export function blockedExtensions(): Set<string> {
    const env = process.env.BLOCKED_EXTENSIONS;
    const list = env !== undefined ? env.split(",") : DEFAULT_BLOCKED;
    return new Set(list.map(s => s.trim().toLowerCase().replace(/^\./, "")).filter(Boolean));
}

export function extensionOf(name: string): string {
    const i = name.lastIndexOf(".");
    return i > 0 ? name.slice(i + 1).toLowerCase() : "";
}

const startsWith = (b: Uint8Array, sig: number[], at = 0) => sig.every((v, i) => b[at + i] === v);
const ascii = (s: string) => Array.from(s, c => c.charCodeAt(0));

const ZIP = [0x50, 0x4b, 0x03, 0x04];
const OLE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

/** Known signatures. Types not listed here are not content-checked (e.g. txt, csv). */
const SIGNATURES: Record<string, (b: Uint8Array) => boolean> = {
    pdf: b => indexOfAscii(b, "%PDF-", 1024) >= 0,
    docx: b => startsWith(b, ZIP), xlsx: b => startsWith(b, ZIP), pptx: b => startsWith(b, ZIP),
    odt: b => startsWith(b, ZIP), ods: b => startsWith(b, ZIP), odp: b => startsWith(b, ZIP),
    doc: b => startsWith(b, OLE), xls: b => startsWith(b, OLE), ppt: b => startsWith(b, OLE), msg: b => startsWith(b, OLE),
    png: b => startsWith(b, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    jpg: b => startsWith(b, [0xff, 0xd8, 0xff]), jpeg: b => startsWith(b, [0xff, 0xd8, 0xff]),
    gif: b => startsWith(b, ascii("GIF8")),
    tif: b => startsWith(b, [0x49, 0x49, 0x2a, 0x00]) || startsWith(b, [0x4d, 0x4d, 0x00, 0x2a]),
    tiff: b => startsWith(b, [0x49, 0x49, 0x2a, 0x00]) || startsWith(b, [0x4d, 0x4d, 0x00, 0x2a]),
    zip: b => startsWith(b, ZIP),
    rtf: b => startsWith(b, ascii("{\\rtf")),
};

function indexOfAscii(b: Uint8Array, needle: string, limit: number): number {
    const n = ascii(needle);
    const max = Math.min(b.length - n.length, limit);
    for (let i = 0; i <= max; i++) if (startsWith(b, n, i)) return i;
    return -1;
}

const isExecutable = (b: Uint8Array) =>
    startsWith(b, ascii("MZ")) ||                       // Windows PE
    startsWith(b, [0x7f, 0x45, 0x4c, 0x46]) ||          // ELF
    startsWith(b, [0xcf, 0xfa, 0xed, 0xfe]) ||          // Mach-O
    startsWith(b, [0xca, 0xfe, 0xba, 0xbe]) ||          // Mach-O fat / Java class
    startsWith(b, ascii("#!"));                         // script with shebang

/**
 * Rejects files that are dangerous regardless of how they're named:
 * blocked extensions (checked on the final extension, so "invoice.pdf.exe" is blocked),
 * executable content under any name, and content that doesn't match a known extension
 * (e.g. an .exe renamed to .pdf).
 */
export function assertFileAllowed(fileName: string, content: Uint8Array): void {
    const ext = extensionOf(fileName);
    if (ext && blockedExtensions().has(ext)) {
        throw new RejectedFileError(`Files of type .${ext} are not allowed`);
    }
    if (isExecutable(content)) {
        throw new RejectedFileError("Executable files are not allowed");
    }
    const check = SIGNATURES[ext];
    if (check && !check(content)) {
        throw new RejectedFileError(`The file content does not match its .${ext} extension`);
    }
}
