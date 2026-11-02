import * as React from "react";
import { Button, Text, makeStyles, tokens } from "@fluentui/react-components";
import { SkeletonPreview } from "./SkeletonCards";
import { DismissRegular, OpenRegular } from "./Icons";
import { IDocumentItem } from "../types";
import { IPreviewTarget } from "../services/DocumentApi";
import { CopyLinkButton } from "./CopyLinkButton";

/** Defence in depth: the backend already restricts hosts; never embed a non-https URL. */
const isHttps = (u: string) => { try { return new URL(u).protocol === "https:"; } catch { return false; } };

const useStyles = makeStyles({
    pane: {
        display: "flex", flexDirection: "column", minWidth: "320px", width: "38%",
        border: `1px solid ${tokens.colorNeutralStroke2}`, borderRadius: tokens.borderRadiusMedium,
        backgroundColor: tokens.colorNeutralBackground2,
    },
    header: {
        display: "flex", alignItems: "center", gap: tokens.spacingHorizontalS,
        padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
        borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    },
    title: { flex: 1, fontWeight: tokens.fontWeightSemibold, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
    frame: { flex: 1, border: "none", minHeight: "560px", backgroundColor: tokens.colorNeutralBackground1 },
    hidden: { display: "none" },
    center: { flex: 1, display: "flex", alignItems: "center", justifyContent: "center", padding: tokens.spacingHorizontalL },
    error: { color: tokens.colorPaletteRedForeground1 },
});

export interface IPreviewPaneProps {
    doc: IDocumentItem;
    target: IPreviewTarget | null;
    error: string | null;
    onClose: () => void;
    onCopyFailed: (doc: IDocumentItem) => void;
}

export const PreviewPane: React.FC<IPreviewPaneProps> = ({ doc, target, error, onClose, onCopyFailed }) => {
    const s = useStyles();
    // Unique per pane so a POST-based preview can target this iframe by name.
    const frameName = React.useMemo(() => `cdm-preview-${Math.random().toString(36).slice(2)}`, []);
    const formRef = React.useRef<HTMLFormElement>(null);
    const isPost = target?.postParameters !== undefined;
    const params = React.useMemo(
        () => (isPost ? Array.from(new URLSearchParams(target!.postParameters).entries()) : []),
        [isPost, target],
    );
    const invalid = !!target && !isHttps(target.url);

    // Graph's postUrl variant needs a form POST into the iframe (the form is rendered by React, inside the control).
    React.useEffect(() => {
        if (isPost && !invalid) formRef.current?.submit();
    }, [isPost, invalid, target]);
    return (
        <aside className={s.pane} aria-label="Document preview">
            <div className={s.header}>
                <Text className={s.title}>Preview: {doc.name}</Text>
                <CopyLinkButton doc={doc} onCopyFailed={onCopyFailed} />
                <Button size="small" appearance="subtle" icon={<OpenRegular />} title="Open in SharePoint"
                        disabled={!doc.webUrl}
                        onClick={() => window.open(doc.webUrl, "_blank", "noopener")} />
                <Button size="small" appearance="subtle" icon={<DismissRegular />} title="Close preview" onClick={onClose} />
            </div>
            {error || invalid ? (
                <div className={s.center}><Text className={s.error}>{error ?? "Preview link was rejected (not https)."}</Text></div>
            ) : target ? (
                <>
                    <iframe className={s.frame} name={frameName} src={isPost ? "about:blank" : target.url}
                            title={`Preview of ${doc.name}`} referrerPolicy="no-referrer"
                            sandbox="allow-scripts allow-same-origin allow-popups allow-forms" />
                    {isPost && (
                        <form ref={formRef} className={s.hidden} method="post" action={target.url} target={frameName}>
                            {params.map(([k, v]) => <input key={k} type="hidden" name={k} value={v} />)}
                        </form>
                    )}
                </>
            ) : (
                <SkeletonPreview />
            )}
        </aside>
    );
};
