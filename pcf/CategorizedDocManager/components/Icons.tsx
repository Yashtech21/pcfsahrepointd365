/*
 * SVG icon set extracted from @fluentui/react-icons (MIT License, (c) Microsoft Corporation).
 * Inlined so the control does not bundle a second copy of Griffel/React alongside the
 * Dataverse-provided React 16 / Fluent 9 platform libraries.
 */
import * as React from "react";

export type IconProps = React.SVGProps<SVGSVGElement>;

function make(name: string, paths: string[]): React.FC<IconProps> {
    const Icon: React.FC<IconProps> = (props) => (
        <svg width="1em" height="1em" viewBox="0 0 20 20" fill="currentColor" aria-hidden="true" focusable="false" {...props}>
            {paths.map((d, i) => <path key={i} d={d} />)}
        </svg>
    );
    Icon.displayName = name;
    return Icon;
}

export const DocumentPdfRegular = make("DocumentPdfRegular", ["M3 12a2 2 0 1 1 0 4H2v1.5a.5.5 0 0 1-1 0v-5c0-.28.22-.5.5-.5zm-1 3h1a1 1 0 1 0 0-2H2zm6-3a2 2 0 0 1 2 2v2a2 2 0 0 1-2 2H6.5a.5.5 0 0 1-.5-.5v-5c0-.28.22-.5.5-.5zm-1 5h1a1 1 0 0 0 1-1v-2a1 1 0 0 0-1-1H7zm7.1-4.99a.5.5 0 0 1 0 .98L14 13h-2v1.5h1.5a.5.5 0 0 1 0 1H12v2a.5.5 0 0 1-1 0v-5c0-.28.22-.5.5-.5H14zM10.6 2c.4 0 .78.16 1.06.44l3.91 3.91c.28.28.44.67.44 1.06V16a2 2 0 0 1-2 2h-1.09q.09-.23.09-.5V17h1a1 1 0 0 0 1-1v-2.39a1.5 1.5 0 0 0 0-2.23V8h-3.5A1.5 1.5 0 0 1 10 6.5V3H6a1 1 0 0 0-1 1v7.63a3 3 0 0 0-1-.52V4c0-1.1.9-2 2-2zM11 6.5c0 .28.22.5.5.5h3.3L11 3.2z"]);
export const DocumentRegular = make("DocumentRegular", ["M6 2a2 2 0 0 0-2 2v12c0 1.1.9 2 2 2h8a2 2 0 0 0 2-2V7.41c0-.4-.16-.78-.44-1.06l-3.91-3.91A1.5 1.5 0 0 0 10.59 2zM5 4a1 1 0 0 1 1-1h4v3.5c0 .83.67 1.5 1.5 1.5H15v8a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1zm9.8 3h-3.3a.5.5 0 0 1-.5-.5V3.2z"]);
export const DocumentTextRegular = make("DocumentTextRegular", ["M6.5 10a.5.5 0 0 0 0 1h7a.5.5 0 0 0 0-1zm0 2a.5.5 0 0 0 0 1h7a.5.5 0 0 0 0-1zm0 2a.5.5 0 0 0 0 1h7a.5.5 0 0 0 0-1zM4 4c0-1.1.9-2 2-2h4.59q.62.01 1.06.44l3.91 3.91c.28.28.44.67.44 1.06V16a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2zm2-1a1 1 0 0 0-1 1v12a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1V8h-3.5A1.5 1.5 0 0 1 10 6.5V3zm5.5 4h3.3L11 3.2v3.3c0 .28.22.5.5.5"]);
export const TableRegular = make("TableRegular", ["M17 5.5A2.5 2.5 0 0 0 14.5 3h-9A2.5 2.5 0 0 0 3 5.5v9A2.5 2.5 0 0 0 5.5 17h9a2.5 2.5 0 0 0 2.5-2.5zm-13 9V13h3v3H5.36A1.5 1.5 0 0 1 4 14.5m8-1.5v3H8v-3zm2.5 3H13v-3h3v1.64A1.5 1.5 0 0 1 14.5 16M12 8v4H8V8zm1 0h3v4h-3zm-1-4v3H8V4zm1 0h1.64A1.5 1.5 0 0 1 16 5.5V7h-3zM7 4v3H4V5.36A1.5 1.5 0 0 1 5.5 4zm0 4v4H4V8z"]);
export const SlideTextRegular = make("SlideTextRegular", ["M5.5 7a.5.5 0 0 0 0 1h4a.5.5 0 0 0 0-1zm0 2.5a.5.5 0 0 0 0 1h7a.5.5 0 0 0 0-1zm-.5 3c0-.28.22-.5.5-.5h5a.5.5 0 0 1 0 1h-5a.5.5 0 0 1-.5-.5M4.5 4A2.5 2.5 0 0 0 2 6.5v7A2.5 2.5 0 0 0 4.5 16h11a2.5 2.5 0 0 0 2.5-2.5v-7A2.5 2.5 0 0 0 15.5 4zM3 6.5C3 5.67 3.67 5 4.5 5h11c.83 0 1.5.67 1.5 1.5v7c0 .83-.67 1.5-1.5 1.5h-11A1.5 1.5 0 0 1 3 13.5z"]);
export const ImageRegular = make("ImageRegular", ["M14 7.5a1.5 1.5 0 1 1-3 0 1.5 1.5 0 0 1 3 0m-1 0a.5.5 0 1 0-1 0 .5.5 0 0 0 1 0M3 6a3 3 0 0 1 3-3h8a3 3 0 0 1 3 3v8a3 3 0 0 1-3 3H6a3 3 0 0 1-3-3zm3-2a2 2 0 0 0-2 2v8q0 .56.28 1.02l4.67-4.59a1.5 1.5 0 0 1 2.1 0l4.67 4.59Q16 14.56 16 14V6a2 2 0 0 0-2-2zm0 12h8a2 2 0 0 0 1.01-.27l-4.66-4.58a.5.5 0 0 0-.7 0l-4.66 4.58A2 2 0 0 0 6 16"]);
export const EyeRegular = make("EyeRegular", ["M3.26 11.6A7 7 0 0 1 10 6c3.2 0 6.06 2.33 6.74 5.6a.5.5 0 0 0 .98-.2A8 8 0 0 0 10 5a8 8 0 0 0-7.72 6.4.5.5 0 0 0 .98.2M10 8a3.5 3.5 0 1 0 0 7 3.5 3.5 0 0 0 0-7m-2.5 3.5a2.5 2.5 0 1 1 5 0 2.5 2.5 0 0 1-5 0"]);
export const ArrowDownloadRegular = make("ArrowDownloadRegular", ["M15.5 17a.5.5 0 0 1 .09 1H4.5a.5.5 0 0 1-.09-1H15.5M10 2a.5.5 0 0 1 .5.41V14.3l3.64-3.65a.5.5 0 0 1 .64-.06l.07.06c.17.17.2.44.06.63l-.06.07-4.5 4.5a.5.5 0 0 1-.25.14L10 16a.5.5 0 0 1-.4-.2l-4.46-4.45a.5.5 0 0 1 .64-.76l.07.06 3.65 3.64V2.5c0-.27.22-.5.5-.5"]);
export const DeleteRegular = make("DeleteRegular", ["M8.5 4h3a1.5 1.5 0 0 0-3 0m-1 0a2.5 2.5 0 0 1 5 0h5a.5.5 0 0 1 0 1h-1.05l-1.2 10.34A3 3 0 0 1 12.27 18H7.73a3 3 0 0 1-2.98-2.66L3.55 5H2.5a.5.5 0 0 1 0-1zM5.74 15.23A2 2 0 0 0 7.73 17h4.54a2 2 0 0 0 1.99-1.77L15.44 5H4.56zM8.5 7.5c.28 0 .5.22.5.5v6a.5.5 0 0 1-1 0V8c0-.28.22-.5.5-.5M12 8a.5.5 0 0 0-1 0v6a.5.5 0 0 0 1 0z"]);
export const FilterRegular = make("FilterRegular", ["M7.5 13h5a.5.5 0 0 1 .09 1H7.5a.5.5 0 0 1-.09-1zh5zm-2-4h9a.5.5 0 0 1 .09 1H5.5a.5.5 0 0 1-.09-1zh9zm-2-4h13a.5.5 0 0 1 .09 1H3.5a.5.5 0 0 1-.09-1zh13z"]);
export const PersonRegular = make("PersonRegular", ["M10 2a4 4 0 1 0 0 8 4 4 0 0 0 0-8M7 6a3 3 0 1 1 6 0 3 3 0 0 1-6 0m-2 5a2 2 0 0 0-2 2c0 1.7.83 2.97 2.13 3.8A9 9 0 0 0 10 18c1.85 0 3.58-.39 4.87-1.2A4.4 4.4 0 0 0 17 13a2 2 0 0 0-2-2zm-1 2a1 1 0 0 1 1-1h10a1 1 0 0 1 1 1c0 1.3-.62 2.28-1.67 2.95A8 8 0 0 1 10 17a8 8 0 0 1-4.33-1.05A3.4 3.4 0 0 1 4 13"]);
export const DismissRegular = make("DismissRegular", ["m4.09 4.22.06-.07a.5.5 0 0 1 .63-.06l.07.06L10 9.29l5.15-5.14a.5.5 0 0 1 .63-.06l.07.06c.18.17.2.44.06.63l-.06.07L10.71 10l5.14 5.15c.18.17.2.44.06.63l-.06.07a.5.5 0 0 1-.63.06l-.07-.06L10 10.71l-5.15 5.14a.5.5 0 0 1-.63.06l-.07-.06a.5.5 0 0 1-.06-.63l.06-.07L9.29 10 4.15 4.85a.5.5 0 0 1-.06-.63l.06-.07z"]);
export const ArrowUploadRegular = make("ArrowUploadRegular", ["M15 3a.5.5 0 0 0 .09-.99H4a.5.5 0 0 0-.09.98L4 3zM9.5 18a.5.5 0 0 0 .5-.41V5.7l3.64 3.65c.17.18.44.2.64.06l.07-.06a.5.5 0 0 0 .06-.63l-.06-.07-4.5-4.5A.5.5 0 0 0 9.6 4h-.1a.5.5 0 0 0-.4.19L4.64 8.65a.5.5 0 0 0 .64.76l.07-.06L9 5.71V17.5c0 .28.22.5.5.5"]);
export const ArrowSyncRegular = make("ArrowSyncRegular", ["M11.41 3.64a.5.5 0 0 0 0-.71L9.3.8a.5.5 0 0 0-.7.7l1 1a7.5 7.5 0 0 0-4.08 13.5.5.5 0 0 0 .6-.8A6.5 6.5 0 0 1 10.14 3.5L8.59 5.04a.5.5 0 0 0 .7.7zM8.6 16.36a.5.5 0 0 0 0 .71l2.12 2.12a.5.5 0 0 0 .7-.7l-1-1a7.5 7.5 0 0 0 4.07-13.5.5.5 0 1 0-.59.8q.37.28.7.61a6.5 6.5 0 0 1-4.73 11.1l1.55-1.55a.5.5 0 1 0-.7-.7z"]);
export const ArrowSortRegular = make("ArrowSortRegular", ["M2.35 7.35 5 4.71V16.5a.5.5 0 0 0 1 0V4.7l2.65 2.65a.5.5 0 0 0 .7-.7l-3.49-3.5A.5.5 0 0 0 5.5 3a.5.5 0 0 0-.39.18L1.65 6.65a.5.5 0 1 0 .7.7m15.3 5.3L15 15.29V3.5a.5.5 0 0 0-1 0v11.8l-2.65-2.65a.5.5 0 0 0-.7.7l3.49 3.5a.5.5 0 0 0 .36.15.5.5 0 0 0 .39-.18l3.46-3.47a.5.5 0 1 0-.7-.7"]);
export const SearchRegular = make("SearchRegular", ["M13.73 14.44a6.5 6.5 0 1 1 .7-.7l3.42 3.4a.5.5 0 0 1-.63.77l-.07-.06zm-.71-.71A5.5 5.5 0 0 0 15 9.5a5.5 5.5 0 1 0-1.98 4.23"]);
export const OpenRegular = make("OpenRegular", ["M6 4a2 2 0 0 0-2 2v8c0 1.1.9 2 2 2h8a2 2 0 0 0 2-2v-2.5a.5.5 0 0 1 1 0V14a3 3 0 0 1-3 3H6a3 3 0 0 1-3-3V6a3 3 0 0 1 3-3h2.5a.5.5 0 0 1 0 1zm5-.5c0-.28.22-.5.5-.5h5c.28 0 .5.22.5.5v5a.5.5 0 0 1-1 0V4.7l-4.15 4.15a.5.5 0 0 1-.7-.7L15.29 4H11.5a.5.5 0 0 1-.5-.5"]);
