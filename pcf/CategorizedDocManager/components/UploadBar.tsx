import * as React from "react";
import {
    Badge, Button, Dropdown, Field, Option, ProgressBar, Text, makeStyles, mergeClasses, shorthands, tokens,
} from "@fluentui/react-components";
import { ArrowUploadRegular, DismissRegular, PersonRegular } from "./Icons";
import { ICategory } from "../categories";
import { formatSize } from "../utils";

/** Overall progress of the current batch, reported by DocumentManager while uploading. */
export interface IUploadProgress {
    current: number;      // 1-based index of the file being sent
    total: number;        // files in the batch
    fileName: string;
    sentBytes: number;    // across the whole batch
    totalBytes: number;
    /** "finishing" = sending is over; refreshing the list. */
    phase: "sending" | "finishing";
    /** Files that uploaded successfully (set in the finishing phase). */
    uploaded?: number;
}

const useStyles = makeStyles({
    root: { display: "flex", gap: tokens.spacingHorizontalL, alignItems: "stretch", flexWrap: "wrap" },
    drop: {
        flex: "1 1 360px", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center",
        minHeight: "64px", padding: tokens.spacingVerticalM, cursor: "pointer",
        border: `1px dashed ${tokens.colorNeutralStroke1}`, borderRadius: tokens.borderRadiusMedium,
        backgroundColor: tokens.colorNeutralBackground2, color: tokens.colorNeutralForeground2,
    },
    dragging: { ...shorthands.borderColor(tokens.colorBrandStroke1), backgroundColor: tokens.colorBrandBackground2 },
    disabled: { cursor: "not-allowed", opacity: 0.6 },
    link: { color: tokens.colorBrandForegroundLink },
    staged: { display: "flex", gap: tokens.spacingHorizontalXS, flexWrap: "wrap", marginTop: tokens.spacingVerticalXS },
    controls: {
        display: "flex", gap: tokens.spacingHorizontalM, alignItems: "flex-end", flexWrap: "wrap",
        padding: tokens.spacingVerticalS, border: `1px solid ${tokens.colorNeutralStroke2}`,
        borderRadius: tokens.borderRadiusMedium,
    },
    contact: { display: "inline-flex", alignItems: "center", gap: "4px", marginBottom: "6px" },
    summary: { marginTop: tokens.spacingVerticalXS, color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 },
    progress: { flexBasis: "100%" },
    refused: { marginTop: tokens.spacingVerticalXS, color: tokens.colorPaletteRedForeground1, fontSize: tokens.fontSizeBase200 },
});

export interface IUploadBarProps {
    categories: ICategory[];
    category: string;
    onCategoryChange: (v: string) => void;
    contactName: string;
    disabledReason?: string;
    busy: boolean;
    progress?: IUploadProgress | null;
    /** Per-table size limit; larger files are refused before upload (the server enforces it as well). */
    maxFileBytes?: number;
    /** Resolves with the files that failed, which stay staged so the user can retry them. */
    onUpload: (files: File[]) => Promise<File[]>;
}

export const UploadBar: React.FC<IUploadBarProps> = (p) => {
    const s = useStyles();
    const inputRef = React.useRef<HTMLInputElement>(null);
    const [dragging, setDragging] = React.useState(false);
    const [staged, setStaged] = React.useState<File[]>([]);
    const [submitting, setSubmitting] = React.useState(false);
    // A ref, not state: it flips synchronously, so a fast double-click can't start a second upload
    // before React re-renders the disabled button.
    const submittingRef = React.useRef(false);
    const busy = p.busy || submitting;
    const disabled = !!p.disabledReason || busy;

    const [refused, setRefused] = React.useState<string[]>([]);

    const addFiles = (list: FileList | null) => {
        if (!list || disabled) return;
        const files = Array.from(list);
        const tooBig = p.maxFileBytes ? files.filter(f => f.size > p.maxFileBytes!) : [];
        const empty = files.filter(f => f.size === 0);
        setRefused([
            ...tooBig.map(f => `${f.name} is ${formatSize(f.size)}; the limit here is ${formatSize(p.maxFileBytes!)}`),
            ...empty.map(f => `${f.name} is empty`),
        ]);
        const ok = files.filter(f => !tooBig.includes(f) && !empty.includes(f));
        setStaged(prev => [...prev, ...ok.filter(f => !prev.some(x => x.name === f.name && x.size === f.size))]);
    };

    const upload = async () => {
        if (submittingRef.current || disabled) return;
        if (!staged.length) { inputRef.current?.click(); return; }
        submittingRef.current = true;
        setSubmitting(true);
        try {
            const failed = await p.onUpload(staged);
            setStaged(failed);
        } finally {
            submittingRef.current = false;
            setSubmitting(false);
        }
    };

    const current = p.categories.find(c => c.value === p.category);
    const stagedBytes = React.useMemo(() => staged.reduce((n, f) => n + f.size, 0), [staged]);
    const pr = p.progress;
    const fraction = pr && pr.totalBytes ? Math.min(pr.sentBytes / pr.totalBytes, 1) : undefined;

    return (
        <div className={s.root}>
            <div
                className={mergeClasses(s.drop, dragging && s.dragging, disabled && s.disabled)}
                role="button" tabIndex={0}
                onClick={() => !disabled && inputRef.current?.click()}
                onKeyDown={e => { if ((e.key === "Enter" || e.key === " ") && !disabled) inputRef.current?.click(); }}
                onDragOver={e => { e.preventDefault(); if (!disabled) setDragging(true); }}
                onDragLeave={() => setDragging(false)}
                onDrop={e => { e.preventDefault(); setDragging(false); addFiles(e.dataTransfer.files); }}
            >
                <Text>
                    {p.disabledReason ?? <>Drag &amp; drop files here or <u className={s.link}>browse</u></>}
                </Text>
                {staged.length > 0 && (
                    <div className={s.staged} onClick={e => e.stopPropagation()}>
                        {staged.map(f => (
                            <Badge key={f.name + f.size} appearance="tint" shape="rounded"
                                   icon={<DismissRegular onClick={() => !busy && setStaged(prev => prev.filter(x => x !== f))} />}
                                   iconPosition="after">
                                {f.name}
                            </Badge>
                        ))}
                    </div>
                )}
                {refused.length > 0 && (
                    <Text className={s.refused} role="alert">Not added: {refused.join("; ")}.</Text>
                )}
                {staged.length > 0 && (
                    <Text className={s.summary} aria-live="polite">
                        {staged.length} file{staged.length === 1 ? "" : "s"} selected · {formatSize(stagedBytes) || "0 B"}
                    </Text>
                )}
                <input ref={inputRef} type="file" multiple hidden
                       onChange={e => { addFiles(e.target.files); e.target.value = ""; }} />
            </div>

            <div className={s.controls}>
                <Field label="Document Category" size="small">
                    <Dropdown
                        size="small"
                        value={current?.label ?? ""}
                        selectedOptions={[p.category]}
                        onOptionSelect={(_, d) => d.optionValue && p.onCategoryChange(d.optionValue)}
                        disabled={disabled}
                    >
                        {p.categories.map(c => <Option key={c.value} value={c.value}>{c.label}</Option>)}
                    </Dropdown>
                </Field>
                <Badge className={s.contact} appearance="tint" shape="rounded" size="large" icon={<PersonRegular />}>
                    Contact: {p.contactName || "—"}
                </Badge>
                <Button appearance="primary" icon={<ArrowUploadRegular />} disabled={disabled} onClick={() => void upload()}>
                    {busy && pr?.phase === "finishing" ? "Finishing…"
                        : busy && pr ? `Uploading ${pr.current}/${pr.total}…` : busy ? "Uploading…"
                        : staged.length ? `Upload (${staged.length})` : "Upload"}
                </Button>
            </div>

            {busy && pr && (
                <Field
                    className={s.progress}
                    size="small"
                    validationState="none"
                    validationMessage={pr.phase === "finishing"
                        ? `${pr.uploaded ?? 0} of ${pr.total} file${pr.total === 1 ? "" : "s"} uploaded · refreshing the document list…`
                        : `Uploading ${pr.current} of ${pr.total}: ${pr.fileName} · `
                            + `${formatSize(pr.sentBytes) || "0 B"} of ${formatSize(pr.totalBytes)}`
                            + (fraction !== undefined ? ` (${Math.round(fraction * 100)}%)` : "")}
                >
                    <ProgressBar value={pr.phase === "finishing" ? undefined : fraction} thickness="large"
                                 aria-label={pr.phase === "finishing" ? "Refreshing documents" : "Upload progress"} />
                </Field>
            )}
        </div>
    );
};
