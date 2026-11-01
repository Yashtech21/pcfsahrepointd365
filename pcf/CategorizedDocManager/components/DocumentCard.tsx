import * as React from "react";
import { Badge, Button, Card, CardFooter, Text, Tooltip, makeStyles, tokens, mergeClasses } from "@fluentui/react-components";
import { ArrowDownloadRegular, DeleteRegular, EyeRegular } from "./Icons";
import { IDocumentItem } from "../types";
import { categoryColor, categoryLabel } from "../categories";
import { formatDate, formatSize } from "../utils";
import { FileTypeIcon } from "./FileTypeIcon";

const useStyles = makeStyles({
    card: { minWidth: "220px", cursor: "default" },
    selected: {
        backgroundColor: tokens.colorBrandBackground2,
        outlineWidth: tokens.strokeWidthThick, outlineStyle: "solid", outlineColor: tokens.colorBrandStroke1,
    },
    head: { display: "flex", gap: tokens.spacingHorizontalS, alignItems: "flex-start" },
    body: { display: "flex", flexDirection: "column", gap: "2px", minWidth: 0 },
    name: { fontWeight: tokens.fontWeightSemibold, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
    meta: { color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 },
    footer: { gap: 0 },
    newBadge: { marginLeft: tokens.spacingHorizontalXS },
    action: { minWidth: "auto", paddingLeft: "4px", paddingRight: "6px" },
});

export interface IDocumentCardProps {
    doc: IDocumentItem;
    sourceLabel: string;
    selected: boolean;
    /** Uploaded during this session: shows a "New" indicator. */
    isNew?: boolean;
    onPreview: (d: IDocumentItem) => void;
    onDownload: (d: IDocumentItem) => void;
    onDelete: (d: IDocumentItem) => void;
}

/**
 * Memoized: with stable callbacks from the parent, toggling the preview or tabs re-renders only the
 * cards whose props changed (e.g. the previously and newly selected card).
 */
export const DocumentCard = React.memo(function DocumentCard({ doc, sourceLabel, selected, isNew, onPreview, onDownload, onDelete }: IDocumentCardProps) {
    const s = useStyles();
    return (
        <Card className={mergeClasses(s.card, selected && s.selected)} size="small">
            <div className={s.head}>
                <FileTypeIcon name={doc.name} />
                <div className={s.body}>
                    <Tooltip content={`${doc.name}${doc.size ? ` (${formatSize(doc.size)})` : ""}`} relationship="label">
                        <Text className={s.name}>{doc.name}</Text>
                    </Tooltip>
                    <div>
                        <Badge appearance="filled" size="medium" style={categoryColor(doc.category)}>
                            {categoryLabel(doc.category) || "Uncategorized"}
                        </Badge>
                        {isNew && <Badge className={s.newBadge} appearance="tint" color="success" size="medium">New</Badge>}
                    </div>
                    <Text className={s.meta}>Source Record: {sourceLabel}</Text>
                    <Text className={s.meta}>
                        Date Uploaded: {formatDate(doc.uploadDate)}
                    </Text>
                </div>
            </div>
            <CardFooter className={s.footer}>
                <Button className={s.action} size="small" appearance="subtle" icon={<EyeRegular />} onClick={() => onPreview(doc)}>Preview</Button>
                <Button className={s.action} size="small" appearance="subtle" icon={<ArrowDownloadRegular />} onClick={() => onDownload(doc)}>Download</Button>
                <Button className={s.action} size="small" appearance="subtle" icon={<DeleteRegular />} onClick={() => onDelete(doc)}>Delete</Button>
            </CardFooter>
        </Card>
    );
});
