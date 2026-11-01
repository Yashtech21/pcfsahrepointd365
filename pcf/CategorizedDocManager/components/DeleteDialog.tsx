import * as React from "react";
import {
    Button, Checkbox, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle,
    Text, makeStyles, tokens,
} from "@fluentui/react-components";
import { IDocumentItem } from "../types";
import { useCategoryCatalog } from "../categories";
import { formatDate } from "../utils";

const useStyles = makeStyles({
    details: {
        display: "grid", gridTemplateColumns: "auto 1fr", columnGap: tokens.spacingHorizontalM, rowGap: "2px",
        margin: `${tokens.spacingVerticalM} 0`, padding: tokens.spacingHorizontalM,
        backgroundColor: tokens.colorNeutralBackground3, borderRadius: tokens.borderRadiusMedium,
    },
    label: { color: tokens.colorNeutralForeground3 },
    value: { fontWeight: tokens.fontWeightSemibold, wordBreak: "break-all" },
    warning: { color: tokens.colorPaletteRedForeground1 },
});

export interface IDeleteDialogProps {
    doc: IDocumentItem | null;
    sourceLabel: string;
    deleting: boolean;
    onCancel: () => void;
    onConfirm: (doc: IDocumentItem) => void;
}

/**
 * Deleting requires an explicit acknowledgement: the user sees exactly which file will be removed
 * and must tick the confirmation box before the Delete button is enabled. The confirmed file name is
 * also sent to the server, which refuses the delete if it doesn't match the item.
 */
export const DeleteDialog: React.FC<IDeleteDialogProps> = ({ doc, sourceLabel, deleting, onCancel, onConfirm }) => {
    const s = useStyles();
    const [acknowledged, setAcknowledged] = React.useState(false);
    const { label: categoryLabel } = useCategoryCatalog();
    React.useEffect(() => setAcknowledged(false), [doc?.id]);

    return (
        <Dialog modalType="alert" open={!!doc} onOpenChange={(_, d) => { if (!d.open && !deleting) onCancel(); }}>
            <DialogSurface>
                <DialogBody>
                    <DialogTitle>Delete this document?</DialogTitle>
                    <DialogContent>
                        <Text>You are about to delete the following file from SharePoint:</Text>
                        {doc && (
                            <div className={s.details}>
                                <Text className={s.label}>File</Text><Text className={s.value}>{doc.name}</Text>
                                <Text className={s.label}>Category</Text><Text>{categoryLabel(doc.category) || "Uncategorized"}</Text>
                                <Text className={s.label}>Source</Text><Text>{sourceLabel}</Text>
                                <Text className={s.label}>Uploaded</Text><Text>{formatDate(doc.uploadDate)}</Text>
                            </div>
                        )}
                        <Text className={s.warning}>
                            It will be removed for everyone with access to this contact's documents and moved to the
                            SharePoint recycle bin.
                        </Text>
                        <div>
                            <Checkbox
                                checked={acknowledged}
                                disabled={deleting}
                                onChange={(_, d) => setAcknowledged(!!d.checked)}
                                label="I understand and want to delete this file"
                            />
                        </div>
                    </DialogContent>
                    <DialogActions>
                        <Button appearance="secondary" disabled={deleting} onClick={onCancel}>Cancel</Button>
                        <Button
                            appearance="primary"
                            disabled={!acknowledged || deleting}
                            onClick={() => doc && onConfirm(doc)}
                        >
                            {deleting ? "Deleting…" : "Delete"}
                        </Button>
                    </DialogActions>
                </DialogBody>
            </DialogSurface>
        </Dialog>
    );
};
