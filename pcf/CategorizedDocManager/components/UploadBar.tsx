import * as React from "react";
import { Badge, Button, Dropdown, Field, Option, Text, makeStyles, mergeClasses, shorthands, tokens } from "@fluentui/react-components";
import { ArrowUploadRegular, DismissRegular, PersonRegular } from "./Icons";
import { ICategory } from "../categories";

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
});

export interface IUploadBarProps {
    categories: ICategory[];
    category: string;
    onCategoryChange: (v: string) => void;
    contactName: string;
    disabledReason?: string;
    busy: boolean;
    onUpload: (files: File[]) => Promise<void>;
}

export const UploadBar: React.FC<IUploadBarProps> = (p) => {
    const s = useStyles();
    const inputRef = React.useRef<HTMLInputElement>(null);
    const [dragging, setDragging] = React.useState(false);
    const [staged, setStaged] = React.useState<File[]>([]);
    const disabled = !!p.disabledReason || p.busy;

    const addFiles = (list: FileList | null) => {
        if (!list || disabled) return;
        const files = Array.from(list);
        setStaged(prev => [...prev, ...files.filter(f => !prev.some(x => x.name === f.name && x.size === f.size))]);
    };

    const upload = async () => {
        if (!staged.length) { inputRef.current?.click(); return; }
        await p.onUpload(staged);
        setStaged([]);
    };

    const current = p.categories.find(c => c.value === p.category);

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
                                   icon={<DismissRegular onClick={() => setStaged(staged.filter(x => x !== f))} />}
                                   iconPosition="after">
                                {f.name}
                            </Badge>
                        ))}
                    </div>
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
                    {p.busy ? "Uploading…" : staged.length ? `Upload (${staged.length})` : "Upload"}
                </Button>
            </div>
        </div>
    );
};
