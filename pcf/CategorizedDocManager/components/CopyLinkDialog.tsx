import * as React from "react";
import {
    Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Input, Text, makeStyles, tokens,
} from "@fluentui/react-components";
import { IDocumentItem } from "../types";
import { copyText, isMacPlatform } from "../utils";

const useStyles = makeStyles({
    body: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS },
    hint: { color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 },
});

/** Fallback when the browser blocks clipboard access: shows the link selected, ready for Ctrl+C. */
export const CopyLinkDialog: React.FC<{ doc: IDocumentItem | null; onClose: () => void }> = ({ doc, onClose }) => {
    const s = useStyles();
    const [status, setStatus] = React.useState("");
    React.useEffect(() => setStatus(""), [doc?.id]);
    const isMac = React.useMemo(() => isMacPlatform(), []);

    return (
        <Dialog open={!!doc} onOpenChange={(_, d) => { if (!d.open) onClose(); }}>
            <DialogSurface>
                <DialogBody>
                    <DialogTitle>Copy link</DialogTitle>
                    <DialogContent className={s.body}>
                        <Text>Your browser didn't allow automatic copying. Press {isMac ? "⌘C" : "Ctrl+C"} to copy the link to <b>{doc?.name}</b>:</Text>
                        <Input
                            readOnly
                            value={doc?.webUrl ?? ""}
                            aria-label={`Link to ${doc?.name ?? "document"}`}
                            autoFocus
                            onFocus={e => e.currentTarget.select()}
                        />
                        <Text className={s.hint}>People who open this link need access to the file in SharePoint.</Text>
                        {status && <Text role="status">{status}</Text>}
                    </DialogContent>
                    <DialogActions>
                        <Button appearance="secondary" onClick={onClose}>Close</Button>
                        <Button appearance="primary" onClick={async () => {
                            if (doc && await copyText(doc.webUrl)) onClose(); else setStatus("Still blocked. Select the link and copy it manually.");
                        }}>Try again</Button>
                    </DialogActions>
                </DialogBody>
            </DialogSurface>
        </Dialog>
    );
};
