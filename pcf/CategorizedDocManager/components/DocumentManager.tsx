import * as React from "react";
import {
    Button, Input, MessageBar, MessageBarActions, MessageBarBody, Tab, TabList, Text, makeStyles, tokens,
} from "@fluentui/react-components";
import { ArrowSortRegular, ArrowSyncRegular, FilterRegular, SearchRegular } from "./Icons";
import { IDocumentItem, IHostContext } from "../types";
import { ICategory, categoryLabel } from "../categories";
import { DocumentApi } from "../services/DocumentApi";
import { prettifyEntityName } from "../utils";
import { UploadBar } from "./UploadBar";
import { DocumentCard } from "./DocumentCard";
import { PreviewPane } from "./PreviewPane";
import { SkeletonCards } from "./SkeletonCards";
import { DeleteDialog } from "./DeleteDialog";

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
});

type Scope = "record" | "all";
type SortKey = "name" | "category" | "date";
interface IPreviewState { doc: IDocumentItem; url: string | null; error: string | null }

export interface IDocumentManagerProps {
    host: IHostContext;
    categories: ICategory[];
    api: DocumentApi | null;
    configError?: string;
    getEntityDisplayName: (logicalName: string) => Promise<string>;
}

export const DocumentManager = React.memo(function DocumentManager(
    { host, categories, api, configError, getEntityDisplayName }: IDocumentManagerProps,
) {
    const s = useStyles();
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
    const [displayNames, setDisplayNames] = React.useState<Record<string, string>>({});

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
    const loadDocs = React.useCallback(async () => {
        const seq = ++requestSeq.current;
        if (!api || !host.contactId) { setDocuments([]); setLoading(false); return; }
        setLoading(true); setLoadError(null);
        try {
            const docs = await api.listByContact(host.contactId);
            if (seq === requestSeq.current) setDocuments(docs);
        } catch (e) {
            if (seq === requestSeq.current) setLoadError(`Could not load documents: ${(e as Error).message}`);
        } finally {
            if (seq === requestSeq.current) setLoading(false);
        }
    }, [api, host.contactId]);

    React.useEffect(() => { void loadDocs(); }, [loadDocs]);

    const sourceLabel = React.useCallback((d: IDocumentItem): string => {
        if (!d.sourceEntity || d.sourceEntity === "contact") return "Contact (Root)";
        const label = displayNames[d.sourceEntity] ?? prettifyEntityName(d.sourceEntity);
        return d.sourceEntity === host.entityName && d.sourceRecordId === host.recordId ? `${label} (this record)` : label;
    }, [displayNames, host.entityName, host.recordId]);

    // Filtering/sorting runs only when its inputs change; opening a preview doesn't redo it.
    const displayed = React.useMemo(() => {
        const q = filter.trim().toLowerCase();
        const list = documents.filter(d => {
            if (scope === "record" && d.sourceRecordId !== host.recordId) return false;
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
    }, [documents, scope, filter, sort, host.recordId, sourceLabel]);

    const toggleSort = React.useCallback((key: SortKey) =>
        setSort(prev => (prev.key === key ? { key, asc: !prev.asc } : { key, asc: key !== "date" })), []);

    const handleUpload = React.useCallback(async (files: File[]): Promise<File[]> => {
        if (!api) return files;
        setUploading(true); setError(null); setInfo(null);
        const failed: File[] = [];
        const messages: string[] = [];
        for (const file of files) {
            try {
                await api.upload({
                    file, category,
                    contactId: host.contactId, contactName: host.contactName,
                    sourceEntity: host.entityName, sourceRecordId: host.recordId,
                });
            } catch (e) {
                failed.push(file);
                messages.push(`${file.name} (${(e as Error).message})`);
            }
        }
        setUploading(false);
        if (failed.length) setError(`Upload failed for ${failed.length} file(s), still selected so you can retry: ${messages.join("; ")}`);
        else setInfo(`${files.length} file(s) uploaded as "${categoryLabel(category)}".`);
        if (failed.length < files.length) await loadDocs();
        return failed;
    }, [api, category, host, loadDocs]);

    const openPreview = React.useCallback(async (doc: IDocumentItem) => {
        if (!api) return;
        setPreview({ doc, url: null, error: null });
        try {
            const url = await api.previewUrl(doc.id);
            setPreview(p => (p && p.doc.id === doc.id ? { ...p, url } : p));
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

    if (configError) {
        return <MessageBar intent="warning"><MessageBarBody>{configError}</MessageBarBody></MessageBar>;
    }

    const uploadDisabledReason = !host.contactId
        ? "Select a contact on this form to manage documents."
        : !host.recordId ? "Save this record before uploading documents." : undefined;
    const sortIcon = (k: SortKey) => (sort.key === k ? (sort.asc ? " ↑" : " ↓") : "");
    const firstLoad = loading && documents.length === 0;

    return (
        <div className={s.root}>
            <Text className={s.title}>Documents</Text>

            <UploadBar
                categories={categories} category={category} onCategoryChange={setCategory}
                contactName={host.contactName} disabledReason={uploadDisabledReason}
                busy={uploading} onUpload={handleUpload}
            />

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
                <Tab value="record" disabled={!host.recordId}>This Record's Documents ({entityLabel})</Tab>
                <Tab value="all">All Contact Documents ({host.contactName || "no contact"})</Tab>
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
                    {filterOpen && (
                        <Input size="small" contentBefore={<SearchRegular />} placeholder="Filter by name, category or source"
                               value={filter} onChange={(_, d) => setFilter(d.value)} autoFocus />
                    )}

                    {firstLoad ? (
                        <SkeletonCards gridClassName={s.grid} />
                    ) : displayed.length === 0 ? (
                        <div className={s.empty}>
                            {!host.contactId ? "No contact selected." : loadError ? "Documents could not be loaded." : "No documents found."}
                        </div>
                    ) : (
                        <div className={s.grid} aria-busy={loading}>
                            {displayed.map(d => (
                                <DocumentCard key={d.id} doc={d} sourceLabel={sourceLabel(d)}
                                              selected={preview?.doc.id === d.id}
                                              onPreview={onPreviewCard}
                                              onDownload={onDownloadCard}
                                              onDelete={onDeleteCard} />
                            ))}
                        </div>
                    )}
                    <Text className={s.footer}>{displayed.length} item(s)</Text>
                </div>

                {preview && (
                    <PreviewPane doc={preview.doc} url={preview.url} error={preview.error} onClose={closePreview} />
                )}
            </div>

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
