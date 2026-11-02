import * as React from "react";
import {
    Button, CounterBadge, Input, MessageBar, MessageBarActions, MessageBarBody, ProgressBar, Tab, TabList, Text,
    ToggleButton, makeStyles, tokens,
} from "@fluentui/react-components";
import { ArrowSortRegular, ArrowSyncRegular, FilterRegular, SearchRegular } from "./Icons";
import { IDocumentItem, IHostContext } from "../types";
import { ICategory, useCategoryCatalog } from "../categories";
import { DocumentApi, IPreviewTarget } from "../services/DocumentApi";
import { TimeoutError } from "../services/http";
import { formatSize, prettifyEntityName } from "../utils";
import { IUploadProgress, UploadBar } from "./UploadBar";
import { DocumentCard } from "./DocumentCard";
import { PreviewPane } from "./PreviewPane";
import { SkeletonCards } from "./SkeletonCards";
import { DeleteDialog } from "./DeleteDialog";
import { CopyLinkDialog } from "./CopyLinkDialog";

const useStyles = makeStyles({
    root: {
        display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM, width: "100%",
        padding: tokens.spacingHorizontalM, boxSizing: "border-box",
        backgroundColor: tokens.colorNeutralBackground1, borderRadius: tokens.borderRadiusMedium,
    },
    title: { fontSize: tokens.fontSizeBase400, fontWeight: tokens.fontWeightSemibold },
    gridHeader: { display: "flex", alignItems: "center", gap: tokens.spacingHorizontalS, flexWrap: "wrap" },
    gridHeaderTitle: { flex: 1, fontWeight: tokens.fontWeightSemibold },
    workspace: { display: "flex", gap: tokens.spacingHorizontalM, alignItems: "stretch" },
    gridCol: { flex: 1, minWidth: 0, display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS },
    grid: { display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(250px, 1fr))", gap: tokens.spacingHorizontalM },
    empty: { padding: tokens.spacingVerticalXXL, textAlign: "center", color: tokens.colorNeutralForeground3 },
    footer: { color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200, textAlign: "right" },
    tabLabel: { display: "inline-flex", alignItems: "center", gap: tokens.spacingHorizontalXS },
    chips: { display: "flex", gap: tokens.spacingHorizontalXS, flexWrap: "wrap", alignItems: "center" },
    chip: { minWidth: "auto" },
    chipCount: { marginLeft: tokens.spacingHorizontalXS },
    // Always rendered with a fixed height so showing/hiding the refresh bar never shifts the layout.
    refreshSlot: { height: "2px", marginTop: `calc(-1 * ${tokens.spacingVerticalXS})` },
    srOnly: {
        position: "absolute", width: "1px", height: "1px", overflow: "hidden",
        clip: "rect(0 0 0 0)", whiteSpace: "nowrap",
    },
});

type Scope = "record" | "all";
type SortKey = "name" | "category" | "date";
interface IPreviewState { doc: IDocumentItem; target: IPreviewTarget | null; error: string | null }

/**
 * After a client-side timeout the server may still have stored the file. Treat a timed-out file as
 * uploaded if a matching document (same size, this record, recent, same base name allowing for
 * SharePoint's " 1" rename suffix) now exists, so a retry doesn't create a duplicate.
 */
function landedDespiteTimeout(file: File, docs: IDocumentItem[], recordId: string, startedAt: number): IDocumentItem | undefined {
    // "Offer Letter.pdf" vs stored "Offer Letter 1.pdf" / sanitised "Offer_Letter.pdf" -> "offerletter"
    const key = (n: string) => n.replace(/\.[^.]+$/, "").replace(/ \d+$/, "").toLowerCase().replace(/[^a-z0-9]/g, "");
    return docs.find(d => d.size === file.size
        && d.sourceRecordId === recordId
        && new Date(d.uploadDate).getTime() >= startedAt - 60000
        && key(d.name) === key(file.name));
}

/** Returns `value` once it has stopped changing for `ms` (used to calm screen-reader announcements). */
function useDebounced<T>(value: T, ms: number): T {
    const [settled, setSettled] = React.useState(value);
    React.useEffect(() => {
        const t = window.setTimeout(() => setSettled(value), ms);
        return () => window.clearTimeout(t);
    }, [value, ms]);
    return settled;
}

export interface IDocumentManagerProps {
    host: IHostContext;
    categories: ICategory[];
    api: DocumentApi;
    getEntityDisplayName: (logicalName: string) => Promise<string>;
    /** Per-table upload limit from configuration (the server enforces it too). */
    maxUploadBytes?: number;
    /** Non-blocking configuration warning shown above the grid. */
    notice?: string;
}

export const DocumentManager = React.memo(function DocumentManager(
    { host, categories, api, getEntityDisplayName, maxUploadBytes, notice }: IDocumentManagerProps,
) {
    const s = useStyles();
    const catalog = useCategoryCatalog();
    const categoryLabel = catalog.label;
    const [scope, setScope] = React.useState<Scope>(host.recordId ? "record" : "all");
    const [documents, setDocuments] = React.useState<IDocumentItem[]>([]);
    const [loading, setLoading] = React.useState(false);
    const [loadError, setLoadError] = React.useState<string | null>(null);
    const [uploading, setUploading] = React.useState(false);
    const [error, setError] = React.useState<string | null>(null);
    const [info, setInfo] = React.useState<string | null>(null);
    const [category, setCategory] = React.useState(categories[0]?.value ?? "");
    const [sort, setSort] = React.useState<{ key: SortKey; asc: boolean }>({ key: "date", asc: false });
    const [filterOpen, setFilterOpen] = React.useState(false);
    const [filter, setFilter] = React.useState("");
    const [preview, setPreview] = React.useState<IPreviewState | null>(null);
    const [toDelete, setToDelete] = React.useState<IDocumentItem | null>(null);
    const [deleting, setDeleting] = React.useState(false);
    const [copyFallback, setCopyFallback] = React.useState<IDocumentItem | null>(null);
    const [displayNames, setDisplayNames] = React.useState<Record<string, string>>({});
    const [uploadProgress, setUploadProgress] = React.useState<IUploadProgress | null>(null);
    const [newIds, setNewIds] = React.useState<ReadonlySet<string>>(() => new Set());
    const [categoryFilter, setCategoryFilter] = React.useState<string | null>(null);
    /** Contact whose documents are in `documents` ("" = none loaded yet). Counts are shown only when it matches. */
    const [loadedFor, setLoadedForState] = React.useState("");
    const loadedForRef = React.useRef("");
    const setLoadedFor = React.useCallback((id: string) => { loadedForRef.current = id; setLoadedForState(id); }, []);

    // Keep the selected category valid when the offered list changes.
    React.useEffect(() => {
        setCategory(c => (categories.some(x => x.value === c) ? c : categories[0]?.value ?? ""));
    }, [categories]);

    // A new record gets an id after its first save: switch to its own documents at that point.
    const prevRecordId = React.useRef(host.recordId);
    React.useEffect(() => {
        if (!prevRecordId.current && host.recordId) setScope("record");
        prevRecordId.current = host.recordId;
    }, [host.recordId]);

    // Resolve display names (from Dataverse metadata) for the host entity and every source entity shown.
    React.useEffect(() => {
        const wanted = new Set([host.entityName, ...documents.map(d => d.sourceEntity)]);
        wanted.delete(""); wanted.delete("contact");
        let cancelled = false;
        wanted.forEach(name => {
            if (displayNames[name]) return;
            void getEntityDisplayName(name).then(label => {
                if (!cancelled) setDisplayNames(prev => (prev[name] === label ? prev : { ...prev, [name]: label }));
            });
        });
        return () => { cancelled = true; };
    }, [host.entityName, documents, getEntityDisplayName]);

    const entityLabel = displayNames[host.entityName] ?? host.entityDisplayName;

    // Only the latest request for the current contact may update state (guards against out-of-order
    // responses when the lookup changes quickly). Identical concurrent requests share one call in DocumentApi.
    const requestSeq = React.useRef(0);
    const loadDocs = React.useCallback(async (): Promise<IDocumentItem[] | null> => {
        const seq = ++requestSeq.current;
        if (!api || !host.contactId) { setDocuments([]); setLoadedFor(""); setLoading(false); return null; }
        const contactId = host.contactId;
        setLoading(true); setLoadError(null);
        // Never show another contact's documents (or counts) while this contact's list is loading or failed.
        if (loadedForRef.current !== contactId) { setDocuments([]); setLoadedFor(""); }
        try {
            const docs = await api.listByContact(contactId);
            if (seq === requestSeq.current) { setDocuments(docs); setLoadedFor(contactId); }
            return docs;
        } catch (e) {
            if (seq === requestSeq.current) setLoadError(`Could not load documents: ${(e as Error).message}`);
        } finally {
            if (seq === requestSeq.current) setLoading(false);
        }
        return null;
    }, [api, host.contactId, setLoadedFor]);

    React.useEffect(() => { void loadDocs(); }, [loadDocs]);

    const sourceLabel = React.useCallback((d: IDocumentItem): string => {
        if (!d.sourceEntity || d.sourceEntity === "contact") return "Contact (Root)";
        const label = displayNames[d.sourceEntity] ?? prettifyEntityName(d.sourceEntity);
        return d.sourceEntity === host.entityName && d.sourceRecordId === host.recordId ? `${label} (this record)` : label;
    }, [displayNames, host.entityName, host.recordId]);

    // "This record" documents, computed once: drives both the tab count and the record-scope list.
    const recordDocs = React.useMemo(
        () => (host.recordId ? documents.filter(d => d.sourceRecordId === host.recordId) : []),
        [documents, host.recordId]);

    // Documents in the active tab, before the category/text filters.
    const scopeDocs = scope === "record" ? recordDocs : documents;
    const countsReady = loadedFor !== "" && loadedFor === host.contactId;

    // Per-category counts within the active tab (for the filter chips), largest first.
    const categoryCounts = React.useMemo(() => {
        const m = new Map<string, number>();
        scopeDocs.forEach(d => m.set(d.category, (m.get(d.category) ?? 0) + 1));
        return Array.from(m.entries()).sort((a, b) => b[1] - a[1] || categoryLabel(a[0]).localeCompare(categoryLabel(b[0])));
    }, [scopeDocs, categoryLabel]);

    // Drop a category filter that no longer matches anything in this tab.
    React.useEffect(() => {
        if (categoryFilter !== null && !categoryCounts.some(([c]) => c === categoryFilter)) setCategoryFilter(null);
    }, [categoryCounts, categoryFilter]);

    // Filtering/sorting runs only when its inputs change; opening a preview doesn't redo it.
    const displayed = React.useMemo(() => {
        const q = filter.trim().toLowerCase();
        const list = scopeDocs.filter(d => {
            if (categoryFilter !== null && d.category !== categoryFilter) return false;
            if (!q) return true;
            return d.name.toLowerCase().includes(q)
                || categoryLabel(d.category).toLowerCase().includes(q)
                || sourceLabel(d).toLowerCase().includes(q);
        });
        const dir = sort.asc ? 1 : -1;
        return list.sort((a, b) => {
            switch (sort.key) {
                case "name": return a.name.localeCompare(b.name) * dir;
                case "category": return categoryLabel(a.category).localeCompare(categoryLabel(b.category)) * dir || a.name.localeCompare(b.name);
                default: return (new Date(a.uploadDate).getTime() - new Date(b.uploadDate).getTime()) * dir;
            }
        });
    }, [scopeDocs, categoryFilter, filter, sort, sourceLabel, categoryLabel]);

    const displayedBytes = React.useMemo(() => displayed.reduce((n, d) => n + (d.size || 0), 0), [displayed]);
    const countText = displayed.length === scopeDocs.length
        ? `${displayed.length} item${displayed.length === 1 ? "" : "s"}`
        : `Showing ${displayed.length} of ${scopeDocs.length} items`;
    const announcedCount = useDebounced(countText, 600);

    const toggleSort = React.useCallback((key: SortKey) =>
        setSort(prev => (prev.key === key ? { key, asc: !prev.asc } : { key, asc: key !== "date" })), []);

    const handleUpload = React.useCallback(async (files: File[]): Promise<File[]> => {
        if (!api) return files;
        setUploading(true); setError(null); setInfo(null);
        const startedAt = Date.now();
        const totalBytes = files.reduce((n, f) => n + f.size, 0);
        let doneBytes = 0;        // bytes of files that actually uploaded
        const uploadedIds: string[] = [];
        const failed: File[] = [];
        const timedOut = new Set<File>();
        const messages = new Map<File, string>();
        for (const [i, file] of files.entries()) {
            const base = doneBytes;
            setUploadProgress({ current: i + 1, total: files.length, fileName: file.name, sentBytes: base, totalBytes, phase: "sending" });
            try {
                const created = await api.upload({
                    file, category,
                    contactId: host.contactId, contactName: host.contactName,
                    sourceEntity: host.entityName, sourceRecordId: host.recordId,
                }, (loaded) => setUploadProgress(p => (p ? { ...p, sentBytes: base + Math.min(loaded, file.size) } : p)));
                if (created?.id) uploadedIds.push(created.id);
                doneBytes += file.size;
            } catch (e) {
                failed.push(file);
                if (e instanceof TimeoutError) timedOut.add(file);
                messages.set(file, `${file.name} (${(e as Error).message})`);
                // Failed bytes don't count: pull the bar back to what really uploaded.
                setUploadProgress(p => (p ? { ...p, sentBytes: doneBytes } : p));
            }
        }

        let stillFailed = failed;
        const okSoFar = files.length - failed.length;
        if (okSoFar > 0 || timedOut.size) {
            setUploadProgress(p => (p ? { ...p, sentBytes: doneBytes, phase: "finishing", uploaded: okSoFar } : p));
            const docs = await loadDocs();
            // A timed-out upload may have completed on the server; don't leave it staged for a duplicate retry.
            if (docs && timedOut.size) {
                stillFailed = failed.filter(f => {
                    if (!timedOut.has(f)) return true;
                    const landed = landedDespiteTimeout(f, docs, host.recordId, startedAt);
                    if (landed) uploadedIds.push(landed.id);
                    return !landed;
                });
            }
        }
        if (uploadedIds.length) setNewIds(prev => new Set([...prev, ...uploadedIds]));
        setUploadProgress(null);
        setUploading(false);
        const succeeded = files.length - stillFailed.length;
        if (stillFailed.length) {
            setError(`Upload failed for ${stillFailed.length} file(s), still selected so you can retry: `
                + stillFailed.map(f => messages.get(f)).join("; "));
            if (succeeded) setInfo(`${succeeded} file(s) uploaded as "${categoryLabel(category)}".`);
        } else {
            setInfo(`${files.length} file(s) uploaded as "${categoryLabel(category)}".`);
        }
        return stillFailed;
    }, [api, category, host, loadDocs, categoryLabel]);

    const openPreview = React.useCallback(async (doc: IDocumentItem) => {
        if (!api) return;
        setPreview({ doc, target: null, error: null });
        try {
            const target = await api.previewTarget(doc.id);
            setPreview(p => (p && p.doc.id === doc.id ? { ...p, target } : p));
        } catch (e) {
            setPreview(p => (p && p.doc.id === doc.id ? { ...p, error: `Preview unavailable: ${(e as Error).message}` } : p));
        }
    }, [api]);

    const download = React.useCallback(async (doc: IDocumentItem) => {
        if (!api) return;
        try {
            const url = await api.downloadUrl(doc.id);
            const a = document.createElement("a");
            a.href = url; a.download = doc.name; a.rel = "noopener";
            document.body.appendChild(a); a.click(); a.remove();
        } catch (e) {
            setError(`Download failed: ${(e as Error).message}`);
        }
    }, [api]);

    const onPreviewCard = React.useCallback((d: IDocumentItem) => void openPreview(d), [openPreview]);
    const onDownloadCard = React.useCallback((d: IDocumentItem) => void download(d), [download]);
    const onDeleteCard = React.useCallback((d: IDocumentItem) => setToDelete(d), []);
    const closePreview = React.useCallback(() => setPreview(null), []);
    const cancelDelete = React.useCallback(() => setToDelete(null), []);
    const onCopyFailed = React.useCallback((d: IDocumentItem) => setCopyFallback(d), []);
    const closeCopyFallback = React.useCallback(() => setCopyFallback(null), []);

    const confirmDelete = React.useCallback(async (doc: IDocumentItem) => {
        if (!api || deleting) return;
        setDeleting(true); setError(null); setInfo(null);
        try {
            await api.remove(doc.id, doc.name);
            setPreview(p => (p?.doc.id === doc.id ? null : p));
            setDocuments(prev => prev.filter(d => d.id !== doc.id));
            setInfo(`"${doc.name}" deleted.`);
        } catch (e) {
            setError(`Delete failed: ${(e as Error).message}`);
        } finally {
            setDeleting(false);
            setToDelete(null);
        }
    }, [api, deleting]);

    const uploadDisabledReason = !host.contactId
        ? "Select a contact on this form to manage documents."
        : !host.recordId ? "Save this record before uploading documents."
        : !categories.length ? "No document categories are enabled for this table. Ask an administrator to add some." : undefined;
    const sortIcon = (k: SortKey) => (sort.key === k ? (sort.asc ? " ↑" : " ↓") : "");
    const firstLoad = loading && documents.length === 0;

    return (
        <div className={s.root}>
            <Text className={s.title}>Documents</Text>

            <UploadBar
                categories={categories} category={category} onCategoryChange={setCategory}
                contactName={host.contactName} disabledReason={uploadDisabledReason}
                busy={uploading} progress={uploadProgress} onUpload={handleUpload}
                maxFileBytes={maxUploadBytes}
            />

            {notice && <MessageBar intent="warning"><MessageBarBody>{notice}</MessageBarBody></MessageBar>}

            {loadError && (
                <MessageBar intent="error">
                    <MessageBarBody>{loadError}</MessageBarBody>
                    <MessageBarActions>
                        <Button size="small" icon={<ArrowSyncRegular />} disabled={loading} onClick={() => void loadDocs()}>Retry</Button>
                    </MessageBarActions>
                </MessageBar>
            )}
            {error && <MessageBar intent="error"><MessageBarBody>{error}</MessageBarBody></MessageBar>}
            {info && <MessageBar intent="success"><MessageBarBody>{info}</MessageBarBody></MessageBar>}

            <TabList selectedValue={scope} onTabSelect={(_, d) => setScope(d.value as Scope)}>
                <Tab value="record" disabled={!host.recordId}
                     aria-label={`This Record's Documents (${entityLabel})${countsReady ? `, ${recordDocs.length} documents` : ""}`}>
                    <span className={s.tabLabel}>
                        This Record's Documents ({entityLabel})
                        {countsReady && host.recordId && <CounterBadge count={recordDocs.length} showZero size="small"
                            appearance={scope === "record" ? "filled" : "ghost"} color={scope === "record" ? "brand" : "informative"} />}
                    </span>
                </Tab>
                <Tab value="all" aria-label={`All Contact Documents (${host.contactName || "no contact"})${countsReady ? `, ${documents.length} documents` : ""}`}>
                    <span className={s.tabLabel}>
                        All Contact Documents ({host.contactName || "no contact"})
                        {countsReady && <CounterBadge count={documents.length} showZero size="small"
                            appearance={scope === "all" ? "filled" : "ghost"} color={scope === "all" ? "brand" : "informative"} />}
                    </span>
                </Tab>
            </TabList>

            <div className={s.workspace}>
                <div className={s.gridCol}>
                    <div className={s.gridHeader}>
                        <Text className={s.gridHeaderTitle}>
                            {scope === "all"
                                ? `Documents related to ${host.contactName || "the contact"} across all records`
                                : `Documents uploaded from this ${entityLabel.toLowerCase()}`}
                        </Text>
                        <Button size="small" appearance="subtle" icon={<ArrowSortRegular />} onClick={() => toggleSort("name")}>File Name{sortIcon("name")}</Button>
                        <Button size="small" appearance="subtle" icon={<ArrowSortRegular />} onClick={() => toggleSort("category")}>Category{sortIcon("category")}</Button>
                        <Button size="small" appearance="subtle" icon={<ArrowSortRegular />} onClick={() => toggleSort("date")}>Date{sortIcon("date")}</Button>
                        <Button size="small" appearance={filterOpen ? "primary" : "subtle"} icon={<FilterRegular />} title="Filter"
                                onClick={() => { setFilterOpen(!filterOpen); if (filterOpen) setFilter(""); }} />
                        <Button size="small" appearance="subtle" icon={<ArrowSyncRegular />} title="Refresh" disabled={loading}
                                onClick={() => void loadDocs()} />
                    </div>
                    <div className={s.refreshSlot}>
                        {loading && !firstLoad && <ProgressBar thickness="medium" aria-label="Refreshing documents" />}
                    </div>

                    {categoryCounts.length > 1 && (
                        <div className={s.chips} role="group" aria-label="Filter by category">
                            <ToggleButton size="small" shape="circular" className={s.chip}
                                          checked={categoryFilter === null} onClick={() => setCategoryFilter(null)}>
                                All<CounterBadge className={s.chipCount} count={scopeDocs.length} showZero size="small" appearance="ghost" />
                            </ToggleButton>
                            {categoryCounts.map(([cat, n]) => (
                                <ToggleButton key={cat || "_none"} size="small" shape="circular" className={s.chip}
                                              checked={categoryFilter === cat}
                                              onClick={() => setCategoryFilter(f => (f === cat ? null : cat))}>
                                    {categoryLabel(cat) || "Uncategorized"}
                                    <CounterBadge className={s.chipCount} count={n} size="small" appearance="ghost" />
                                </ToggleButton>
                            ))}
                        </div>
                    )}

                    {filterOpen && (
                        <Input size="small" contentBefore={<SearchRegular />} placeholder="Filter by name, category or source"
                               value={filter} onChange={(_, d) => setFilter(d.value)} autoFocus />
                    )}

                    {firstLoad ? (
                        <SkeletonCards gridClassName={s.grid} />
                    ) : displayed.length === 0 ? (
                        <div className={s.empty}>
                            {!host.contactId ? "No contact selected."
                                : loadError ? "Documents could not be loaded."
                                : scopeDocs.length ? "No documents match the current filters."
                                : "No documents found."}
                        </div>
                    ) : (
                        <div className={s.grid} aria-busy={loading}>
                            {displayed.map(d => (
                                <DocumentCard key={d.id} doc={d} sourceLabel={sourceLabel(d)}
                                              selected={preview?.doc.id === d.id}
                                              isNew={newIds.has(d.id)}
                                              onPreview={onPreviewCard}
                                              onDownload={onDownloadCard}
                                              onDelete={onDeleteCard}
                                              onCopyFailed={onCopyFailed} />
                            ))}
                        </div>
                    )}
                    {!firstLoad && (
                        <Text className={s.footer}>
                            {countText}{displayedBytes ? ` · ${formatSize(displayedBytes)}` : ""}
                        </Text>
                    )}
                    {/* Announced once filtering settles, not on every keystroke. */}
                    <div className={s.srOnly} aria-live="polite" aria-atomic="true">{firstLoad ? "" : announcedCount}</div>
                </div>

                {preview && (
                    <PreviewPane doc={preview.doc} target={preview.target} error={preview.error} onClose={closePreview}
                                 onCopyFailed={onCopyFailed} />
                )}
            </div>

            <CopyLinkDialog doc={copyFallback} onClose={closeCopyFallback} />

            <DeleteDialog
                doc={toDelete}
                sourceLabel={toDelete ? sourceLabel(toDelete) : ""}
                deleting={deleting}
                onCancel={cancelDelete}
                onConfirm={d => void confirmDelete(d)}
            />
        </div>
    );
});
