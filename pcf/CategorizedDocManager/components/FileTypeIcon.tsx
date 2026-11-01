import * as React from "react";
import {
    DocumentPdfRegular, DocumentRegular, ImageRegular, TableRegular, SlideTextRegular, DocumentTextRegular,
} from "./Icons";
import { fileExtension } from "../utils";

const COLORS: Record<string, string> = {
    pdf: "#d13438", doc: "#2b579a", docx: "#2b579a", xls: "#217346", xlsx: "#217346", csv: "#217346",
    ppt: "#b7472a", pptx: "#b7472a",
};

export const FileTypeIcon: React.FC<{ name: string; size?: number }> = ({ name, size = 32 }) => {
    const ext = fileExtension(name);
    const style: React.CSSProperties = { fontSize: size, color: COLORS[ext] ?? "#605e5c", flexShrink: 0 };
    switch (ext) {
        case "pdf": return <DocumentPdfRegular style={style} />;
        case "doc": case "docx": return <DocumentTextRegular style={style} />;
        case "xls": case "xlsx": case "csv": return <TableRegular style={style} />;
        case "ppt": case "pptx": return <SlideTextRegular style={style} />;
        case "png": case "jpg": case "jpeg": case "gif": case "bmp": return <ImageRegular style={style} />;
        default: return <DocumentRegular style={style} />;
    }
};
