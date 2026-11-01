import * as React from "react";
import { FluentProvider, webLightTheme, Theme } from "@fluentui/react-components";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { DocumentManager } from "./components/DocumentManager";
import { categoriesFor, ICategory } from "./categories";
import { AuthService } from "./services/AuthService";
import { DocumentApi } from "./services/DocumentApi";
import { IHostContext } from "./types";
import { normalizeGuid, prettifyEntityName } from "./utils";

interface IContextInfo { entityId?: string; entityTypeName?: string }

/**
 * Dynamics calls updateView many times while a form loads and on every field change. Everything
 * handed to React here is cached and only replaced when its inputs actually change, so the
 * component tree sees stable props and doesn't refetch or re-render needlessly.
 */
export class CategorizedDocManager implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private context!: ComponentFramework.Context<IInputs>;
    private api: DocumentApi | null = null;
    private apiKey = "";
    private categories: ICategory[] = [];
    private categoriesKey = "";
    private host: IHostContext | null = null;
    private userPrincipalName!: Promise<string | undefined>;
    private displayNames = new Map<string, Promise<string>>();

    public init(context: ComponentFramework.Context<IInputs>): void {
        this.context = context;
        // The user's UPN is an MSAL login hint so silent SSO picks the right account. Resolved once.
        this.userPrincipalName = context.webAPI
            .retrieveRecord("systemuser", context.userSettings.userId.replace(/[{}]/g, ""), "?$select=domainname")
            .then(u => (u as { domainname?: string }).domainname || undefined)
            .catch(() => undefined);
    }

    /** Cached per entity; passed to React, which stores the result in state (re-renders itself). */
    private getEntityDisplayName = (logicalName: string): Promise<string> => {
        let p = this.displayNames.get(logicalName);
        if (!p) {
            p = this.context.utils.getEntityMetadata(logicalName)
                .then(m => (m as { DisplayName?: string }).DisplayName || prettifyEntityName(logicalName))
                .catch(() => prettifyEntityName(logicalName));
            this.displayNames.set(logicalName, p);
        }
        return p;
    };

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        this.context = context;
        const p = context.parameters;
        const { entityId, entityTypeName } = (context.mode as unknown as { contextInfo?: IContextInfo }).contextInfo ?? {};
        const entityName = (entityTypeName ?? "").toLowerCase();

        const lookup = p.contactLookupField.raw?.[0];
        const next: IHostContext = {
            entityName,
            entityDisplayName: prettifyEntityName(entityName) || "Record",
            recordId: normalizeGuid(entityId),
            contactId: normalizeGuid(lookup?.id),
            contactName: lookup?.name ?? "",
        };
        const h = this.host;
        if (!h || h.entityName !== next.entityName || h.recordId !== next.recordId
            || h.contactId !== next.contactId || h.contactName !== next.contactName) {
            this.host = next;
        }

        const catKey = `${entityName}|${p.categoryConfig.raw ?? ""}`;
        if (catKey !== this.categoriesKey) {
            this.categoriesKey = catKey;
            this.categories = categoriesFor(entityName, p.categoryConfig.raw);
        }

        const cfg = {
            apiBaseUrl: (p.apiBaseUrl.raw ?? "").trim(),
            tenantId: (p.tenantId.raw ?? "").trim(),
            clientId: (p.clientId.raw ?? "").trim(),
            apiScope: (p.apiScope.raw ?? "").trim(),
        };
        const missing = Object.entries(cfg).filter(([, v]) => !v).map(([k]) => k);
        let configError: string | undefined;
        if (missing.length) {
            configError = `Categorized Document Manager is not configured. Missing: ${missing.join(", ")}.`;
            this.api = null;
            this.apiKey = "";
        } else {
            const key = JSON.stringify(cfg);
            if (key !== this.apiKey) {
                this.apiKey = key;
                this.api = new DocumentApi(cfg.apiBaseUrl,
                    new AuthService(cfg.tenantId, cfg.clientId, cfg.apiScope, this.userPrincipalName));
            }
        }

        const theme = (context.fluentDesignLanguage?.tokenTheme as Theme | undefined) ?? webLightTheme;
        return React.createElement(
            FluentProvider,
            { theme, className: "categorized-doc-manager" },
            React.createElement(DocumentManager, {
                host: this.host!,
                categories: this.categories,
                api: this.api,
                configError,
                getEntityDisplayName: this.getEntityDisplayName,
            }),
        );
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
        this.api = null;
    }
}
