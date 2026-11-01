import * as React from "react";
import {
    Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Input,
    MessageBar, MessageBarBody, Spinner, Tab, TabList, Text, makeStyles, tokens,
} from "@fluentui/react-components";
import { ArrowSortRegular, ArrowSyncRegular, FilterRegular, SearchRegular } from "./Icons";
import { IDocumentItem, IHostContext } from "../types";
import { ICategory, categoryLabel } from "../categories";
import { DocumentApi } from "../services/DocumentApi";
import { prettifyEntityName } from "../utils";
import { UploadBar } from "./UploadBar";
import { DocumentCard } from "./DocumentCard";
import { PreviewPane } from "./PreviewPane";

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

export interface IDocumentManagerProps {
    host: IHostContext;
    categories: ICategory[];
    api: DocumentApi | null;
    configError?: string;
}

export const DocumentManager: React.FC<IDocumentManagerProps> = ({ host, categories, api, configError }) => {
    const s = useStyles();
    const [scope, setScope] = React.useState<Scope>(host.recordId ? "record" : "all");
    const [documents, setDocuments] = React.useState<IDocumentItem[]>([]);
    const [loading, setLoading] = React.useState(false);
    const [uploading, setUploading] = React.useState(false);
    const [error, setError] = React.useState<string | null>(null);
    const [info, setInfo] = React.useState<string | null>(null);
    const [category, setCategory] = React.useState(categories[0]?.value ?? "");
    const [sort, setSort] = React.useState<{ key: SortKey; asc: boolean }>({ key: "date", asc: false });
    const [filterOpen, setFilterOpen] = React.useState(false);
    const [filter, setFilter] = React.useState("");
    const [preview, setPreview] = React.useState<{ doc: IDocumentItem; url: string | null; error: string | null } | null>(null);
    const [toDelete, setToDelete] = React.useState<IDocumentItem | null>(null);

    React.useEffect(() => {
        if (!categories.some(c => c.value === category)) setCategory(categories[0]?.value ?? "");
    }, [categories]);

    const loadDocs = React.useCallback(async () => {
        if (!api || !host.contactId) { setDocuments([]); return; }
        setLoading(true); setError(null);
        try {
            setDocuments(await api.listByContact(host.contactId));
        } catch (e) {
            setError(`Could not load documents: ${(e as Error).message}`);
        } finally {
            setLoading(false);
        }
    }, [api, host.contactId]);

    React.useEffect(() => { void loadDocs(); }, [loadDocs]);

    const sourceLabel = (d: IDocumentItem): string => {
        if (!d.sourceEntity || d.sourceEntity === "contact") return "Contact (Root)";
        if (d.sourceEntity === host.entityName) {
            return d.sourceRecordId === host.recordId ? `${host.entityDisplayName} (this record)` : host.entityDisplayName;
        }
        return prettifyEntityName(d.sourceEntity);
    };

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
    }, [documents, scope, filter, sort, host.recordId, host.entityName]);

    const toggleSort = (key: SortKey) =>
        setSort(prev => (prev.key === key ? { key, asc: !prev.asc } : { key, asc: key !== "date" }));

    const handleUpload = async (files: File[]) => {
        if (!api) return;
        setUploading(true); setError(null); setInfo(null);
        const failed: string[] = [];
        for (const file of files) {
            try {
                await api.upload({
                    file, category,
                    contactId: host.contactId, contactName: host.contactName,
                    sourceEntity: host.entityName, sourceRecordId: host.recordId,
                });
            } catch (e) {
                failed.push(`${file.name} (${(e as Error).message})`);
            }
        }
        setUploading(false);
        if (failed.length) setError(`Upload failed: ${failed.join("; ")}`);
        else setInfo(`${files.length} file(s) uploaded as "${categoryLabel(category)}".`);
        await loadDocs();
    };

    const openPreview = async (doc: IDocumentItem) => {
        if (!api) return;
        setPreview({ doc, url: null, error: null });
        try {
            const url = await api.previewUrl(doc.id);
            setPreview(p => (p && p.doc.id === doc.id ? { ...p, url } : p));
        } catch (e) {
            setPreview(p => (p && p.doc.id === doc.id ? { ...p, error: `Preview unavailable: ${(e as Error).message}` } : p));
        }
    };

    const download = async (doc: IDocumentItem) => {
        if (!api) return;
        try {
            const url = await api.downloadUrl(doc.id);
            const a = document.createElement("a");
            a.href = url; a.download = doc.name; a.rel = "noopener";
            document.body.appendChild(a); a.click(); a.remove();
        } catch (e) {
            setError(`Download failed: ${(e as Error).message}`);
        }
    };

    const confirmDelete = async () => {
        if (!api || !toDelete) return;
        const doc = toDelete;
        setToDelete(null);
        try {
            await api.remove(doc.id);
            if (preview?.doc.id === doc.id) setPreview(null);
            setInfo(`"${doc.name}" deleted.`);
            await loadDocs();
        } catch (e) {
            setError(`Delete failed: ${(e as Error).message}`);
        }
    };

    if (configError) {
        return <MessageBar intent="warning"><MessageBarBody>{configError}</MessageBarBody></MessageBar>;
    }

    const uploadDisabledReason = !host.contactId
        ? "Select a contact on this form to manage documents."
        : !host.recordId ? "Save this record before uploading documents." : undefined;
    const sortIcon = (k: SortKey) => (sort.key === k ? (sort.asc ? " ↑" : " ↓") : "");

    return (
        <div className={s.root}>
            <Text className={s.title}>Documents</Text>

            <UploadBar
                categories={categories} category={category} onCategoryChange={setCategory}
                contactName={host.contactName} disabledReason={uploadDisabledReason}
                busy={uploading} onUpload={handleUpload}
            />

            {error && <MessageBar intent="error"><MessageBarBody>{error}</MessageBarBody></MessageBar>}
            {info && <MessageBar intent="success"><MessageBarBody>{info}</MessageBarBody></MessageBar>}

            <TabList selectedValue={scope} onTabSelect={(_, d) => setScope(d.value as Scope)}>
                <Tab value="record" disabled={!host.recordId}>This Record's Documents ({host.entityDisplayName})</Tab>
                <Tab value="all">All Contact Documents ({host.contactName || "no contact"})</Tab>
            </TabList>

            <div className={s.workspace}>
                <div className={s.gridCol}>
                    <div className={s.gridHeader}>
                        <Text className={s.gridHeaderTitle}>
                            {scope === "all"
                                ? `Documents related to ${host.contactName || "the contact"} across all records`
                                : `Documents uploaded from this ${host.entityDisplayName.toLowerCase()}`}
                        </Text>
                        <Button size="small" appearance="subtle" icon={<ArrowSortRegular />} onClick={() => toggleSort("name")}>File Name{sortIcon("name")}</Button>
                        <Button size="small" appearance="subtle" icon={<ArrowSortRegular />} onClick={() => toggleSort("category")}>Category{sortIcon("category")}</Button>
                        <Button size="small" appearance="subtle" icon={<ArrowSortRegular />} onClick={() => toggleSort("date")}>Date{sortIcon("date")}</Button>
                        <Button size="small" appearance={filterOpen ? "primary" : "subtle"} icon={<FilterRegular />} title="Filter"
                                onClick={() => { setFilterOpen(!filterOpen); if (filterOpen) setFilter(""); }} />
                        <Button size="small" appearance="subtle" icon={<ArrowSyncRegular />} title="Refresh" onClick={() => void loadDocs()} />
                    </div>
                    {filterOpen && (
                        <Input size="small" contentBefore={<SearchRegular />} placeholder="Filter by name, category or source"
                               value={filter} onChange={(_, d) => setFilter(d.value)} autoFocus />
                    )}

                    {loading ? (
                        <Spinner label="Loading documents…" />
                    ) : displayed.length === 0 ? (
                        <div className={s.empty}>
                            {host.contactId ? "No documents found." : "No contact selected."}
                        </div>
                    ) : (
                        <div className={s.grid}>
                            {displayed.map(d => (
                                <DocumentCard key={d.id} doc={d} sourceLabel={sourceLabel(d)}
                                              selected={preview?.doc.id === d.id}
                                              onPreview={doc => void openPreview(doc)}
                                              onDownload={doc => void download(doc)}
                                              onDelete={setToDelete} />
                            ))}
                        </div>
                    )}
                    <Text className={s.footer}>{displayed.length} item(s)</Text>
                </div>

                {preview && (
                    <PreviewPane doc={preview.doc} url={preview.url} error={preview.error} onClose={() => setPreview(null)} />
                )}
            </div>

            <Dialog open={!!toDelete} onOpenChange={(_, d) => { if (!d.open) setToDelete(null); }}>
                <DialogSurface>
                    <DialogBody>
                        <DialogTitle>Delete document?</DialogTitle>
                        <DialogContent>
                            “{toDelete?.name}” will be moved to the SharePoint recycle bin.
                        </DialogContent>
                        <DialogActions>
                            <Button appearance="secondary" onClick={() => setToDelete(null)}>Cancel</Button>
                            <Button appearance="primary" onClick={() => void confirmDelete()}>Delete</Button>
                        </DialogActions>
                    </DialogBody>
                </DialogSurface>
            </Dialog>
        </div>
    );
};
