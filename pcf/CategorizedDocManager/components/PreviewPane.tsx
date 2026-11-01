import * as React from "react";
import { Button, Spinner, Text, makeStyles, tokens } from "@fluentui/react-components";
import { DismissRegular, OpenRegular } from "./Icons";
import { IDocumentItem } from "../types";

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
    frame: { flex: 1, border: "none", minHeight: "560px", backgroundColor: "#fff" },
    center: { flex: 1, display: "flex", alignItems: "center", justifyContent: "center", padding: tokens.spacingHorizontalL },
    error: { color: tokens.colorPaletteRedForeground1 },
});

export interface IPreviewPaneProps {
    doc: IDocumentItem;
    url: string | null;
    error: string | null;
    onClose: () => void;
}

export const PreviewPane: React.FC<IPreviewPaneProps> = ({ doc, url, error, onClose }) => {
    const s = useStyles();
    return (
        <aside className={s.pane} aria-label="Document preview">
            <div className={s.header}>
                <Text className={s.title}>Preview: {doc.name}</Text>
                <Button size="small" appearance="subtle" icon={<OpenRegular />} title="Open in SharePoint"
                        onClick={() => window.open(doc.webUrl, "_blank", "noopener")} />
                <Button size="small" appearance="subtle" icon={<DismissRegular />} title="Close preview" onClick={onClose} />
            </div>
            {error ? (
                <div className={s.center}><Text className={s.error}>{error}</Text></div>
            ) : url ? (
                <iframe className={s.frame} src={url} title={`Preview of ${doc.name}`}
                        sandbox="allow-scripts allow-same-origin allow-popups allow-forms" />
            ) : (
                <div className={s.center}><Spinner label="Loading preview…" /></div>
            )}
        </aside>
    );
};
