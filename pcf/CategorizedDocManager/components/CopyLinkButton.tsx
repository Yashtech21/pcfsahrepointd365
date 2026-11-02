import * as React from "react";
import { Button, Tooltip } from "@fluentui/react-components";
import { CheckmarkRegular, LinkRegular } from "./Icons";
import { IDocumentItem } from "../types";
import { copyText } from "../utils";

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
    const [copied, setCopied] = React.useState(false);
    const timer = React.useRef<number>();
    React.useEffect(() => () => window.clearTimeout(timer.current), []);

    const copy = async () => {
        if (!doc.webUrl) return;
        if (await copyText(doc.webUrl)) {
            setCopied(true);
            window.clearTimeout(timer.current);
            timer.current = window.setTimeout(() => setCopied(false), 2000);
        } else {
            onCopyFailed(doc);
        }
    };

    const label = !doc.webUrl ? "Link not available" : copied ? "Link copied" : "Copy link";
    return (
        <Tooltip content={label} relationship="label">
            <Button
                className={className}
                size="small"
                appearance="subtle"
                icon={copied ? <CheckmarkRegular /> : <LinkRegular />}
                disabled={!doc.webUrl}
                onClick={() => void copy()}
            />
        </Tooltip>
    );
};
