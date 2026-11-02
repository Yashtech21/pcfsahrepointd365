import * as React from "react";
import { Button, Tooltip, makeStyles } from "@fluentui/react-components";
import { CheckmarkRegular, LinkRegular } from "./Icons";
import { IDocumentItem } from "../types";
import { copyText } from "../utils";

const useStyles = makeStyles({
    srOnly: { position: "absolute", width: "1px", height: "1px", overflow: "hidden", clip: "rect(0 0 0 0)", whiteSpace: "nowrap" },
});

export interface ICopyLinkButtonProps {
    doc: IDocumentItem;
    /** Called when the browser blocks clipboard access, so the link can be shown for manual copying. */
    onCopyFailed: (doc: IDocumentItem) => void;
    className?: string;
}

/**
 * Copies the document's SharePoint link (the backend has already checked it's a trusted https link).
 * Shows a checkmark for 2 seconds after copying. People who open the link need SharePoint access to the file.
 */
export const CopyLinkButton: React.FC<ICopyLinkButtonProps> = ({ doc, onCopyFailed, className }) => {
    const s = useStyles();
    const [copied, setCopied] = React.useState(false);
    const timer = React.useRef<number>();
    // The card can unmount while the clipboard promise is pending (list refresh, filter): don't touch state then.
    const mounted = React.useRef(true);
    React.useEffect(() => () => { mounted.current = false; window.clearTimeout(timer.current); }, []);

    const copy = async () => {
        if (!doc.webUrl) return;
        const ok = await copyText(doc.webUrl);
        if (!mounted.current) return;
        if (ok) {
            setCopied(true);
            window.clearTimeout(timer.current);
            timer.current = window.setTimeout(() => setCopied(false), 2000);
        } else {
            onCopyFailed(doc);
        }
    };

    const label = !doc.webUrl ? "Link not available" : copied ? "Link copied" : "Copy link";
    return (
        <>
            <Tooltip content={label} relationship="label">
                <Button
                    className={className}
                    size="small"
                    appearance="subtle"
                    icon={copied ? <CheckmarkRegular /> : <LinkRegular />}
                    // disabledFocusable keeps hover/focus, so the "Link not available" tooltip can still explain why.
                    disabledFocusable={!doc.webUrl}
                    onClick={() => void copy()}
                />
            </Tooltip>
            {/* Screen readers don't announce aria-label changes: say it explicitly. */}
            <span className={s.srOnly} role="status" aria-live="polite">{copied ? `Link to ${doc.name} copied` : ""}</span>
        </>
    );
};
