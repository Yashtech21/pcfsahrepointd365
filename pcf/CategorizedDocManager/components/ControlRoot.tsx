import * as React from "react";
import { Button, MessageBar, MessageBarActions, MessageBarBody, MessageBarTitle, makeStyles, tokens } from "@fluentui/react-components";
import { ArrowSyncRegular } from "./Icons";
import { CategoryCatalogContext, ICategory, createCatalog } from "../categories";
import { AuthService } from "../services/AuthService";
import { DocumentApi } from "../services/DocumentApi";
import { IApiSettings, ITableSettings, SettingsService } from "../services/SettingsService";
import { IHostContext } from "../types";
import { DocumentManager } from "./DocumentManager";
import { SkeletonCards } from "./SkeletonCards";

const useStyles = makeStyles({
    root: { padding: tokens.spacingHorizontalM },
    grid: { display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(250px, 1fr))", gap: tokens.spacingHorizontalM },
});

export interface IControlServices {
    settings: SettingsService;
    getEntityDisplayName: (logicalName: string) => Promise<string>;
    userPrincipalName: Promise<string | undefined>;
}

export interface IControlRootProps {
    host: IHostContext;
    services: IControlServices;
    /** Per-form property values; any non-empty one overrides the matching environment variable. */
    overrides: Partial<IApiSettings>;
    /** Logical name of the lookup column the control is bound to on this form. */
    boundLookupColumn: string;
}

type State =
    | { status: "loading" }
    | { status: "error"; message: string }
    | { status: "unconfigured" }
    | { status: "ready"; api: IApiSettings; table: ITableSettings; catalog: ICategory[] };

/**
 * Loads configuration (environment variables, table configuration, category catalog) once, then renders
 * the document manager. All "is this set up?" problems surface here as clear messages, not broken UI.
 */
export const ControlRoot = React.memo(function ControlRoot({ host, services, overrides, boundLookupColumn }: IControlRootProps) {
    const s = useStyles();
    const [state, setState] = React.useState<State>({ status: "loading" });
    const [attempt, setAttempt] = React.useState(0);

    React.useEffect(() => {
        let cancelled = false;
        setState({ status: "loading" });
        Promise.all([
            services.settings.loadApiSettings(overrides),
            services.settings.loadTable(host.entityName),
            services.settings.loadCatalog(),
        ]).then(([api, table, catalog]) => {
            if (cancelled) return;
            setState(table ? { status: "ready", api, table, catalog } : { status: "unconfigured" });
        }).catch(e => {
            if (!cancelled) setState({ status: "error", message: (e as Error).message });
        });
        return () => { cancelled = true; };
    }, [services, overrides, host.entityName, attempt]);

    const ready = state.status === "ready" ? state : null;
    const apiKey = ready ? JSON.stringify(ready.api) : "";
    const api = React.useMemo(() => {
        if (!ready) return null;
        const a = ready.api;
        return new DocumentApi(a.apiBaseUrl, new AuthService(a.tenantId, a.clientId, a.apiScope, services.userPrincipalName));
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [apiKey, services.userPrincipalName]);

    const catalog = React.useMemo(
        () => createCatalog(ready ? [...ready.catalog, ...ready.table.categories] : []),
        [ready?.catalog, ready?.table]);

    if (state.status === "loading") {
        return <div className={s.root}><SkeletonCards gridClassName={s.grid} count={3} /></div>;
    }
    if (state.status === "error") {
        return (
            <MessageBar intent="error">
                <MessageBarBody><MessageBarTitle>Document Manager</MessageBarTitle>{state.message}</MessageBarBody>
                <MessageBarActions>
                    <Button size="small" icon={<ArrowSyncRegular />} onClick={() => setAttempt(n => n + 1)}>Retry</Button>
                </MessageBarActions>
            </MessageBar>
        );
    }
    if (state.status === "unconfigured") {
        return (
            <MessageBar intent="warning">
                <MessageBarBody>
                    <MessageBarTitle>Not set up for this table</MessageBarTitle>
                    Document management isn't configured for <b>{host.entityDisplayName}</b> ({host.entityName}). An administrator can
                    add an active "Document Manager Configuration" row with Table Logical Name <b>{host.entityName}</b>.
                </MessageBarBody>
            </MessageBar>
        );
    }

    const { table } = state;
    const notice = boundLookupColumn && table.contactLookupColumn && boundLookupColumn !== table.contactLookupColumn
        ? `This control is bound to "${boundLookupColumn}", but the configuration for ${table.name} says the contact lookup is `
          + `"${table.contactLookupColumn}". Update the configuration row or the form so they match.`
        : undefined;

    return (
        <CategoryCatalogContext.Provider value={catalog}>
            <DocumentManager
                host={host}
                categories={table.categories}
                api={api!}
                getEntityDisplayName={services.getEntityDisplayName}
                maxUploadBytes={table.maxUploadMb ? table.maxUploadMb * 1024 * 1024 : undefined}
                notice={notice}
            />
        </CategoryCatalogContext.Provider>
    );
});
