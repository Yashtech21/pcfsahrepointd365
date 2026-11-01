import * as React from "react";
import { FluentProvider, webLightTheme, Theme } from "@fluentui/react-components";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { DocumentManager } from "./components/DocumentManager";
import { categoriesFor } from "./categories";
import { AuthService } from "./services/AuthService";
import { DocumentApi } from "./services/DocumentApi";
import { IHostContext } from "./types";
import { normalizeGuid, prettifyEntityName } from "./utils";

interface IContextInfo { entityId?: string; entityTypeName?: string }

export class CategorizedDocManager implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private context!: ComponentFramework.Context<IInputs>;
    private notifyOutputChanged!: () => void;
    private api: DocumentApi | null = null;
    private apiKey = "";
    private entityDisplayName = "";
    private userPrincipalName: string | undefined;

    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void): void {
        this.context = context;
        this.notifyOutputChanged = notifyOutputChanged;

        const { entityTypeName } = this.contextInfo(context);
        if (entityTypeName) {
            context.utils.getEntityMetadata(entityTypeName)
                .then(m => { this.entityDisplayName = (m as { DisplayName?: string }).DisplayName ?? ""; this.notifyOutputChanged(); })
                .catch(() => { /* fall back to prettified logical name */ });
        }
        // The user's UPN is used as an MSAL login hint so silent SSO picks the right account.
        context.webAPI.retrieveRecord("systemuser", context.userSettings.userId.replace(/[{}]/g, ""), "?$select=domainname")
            .then(u => { this.userPrincipalName = (u as { domainname?: string }).domainname; })
            .catch(() => { /* login hint is optional */ });
    }

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        this.context = context;
        const p = context.parameters;
        const { entityId, entityTypeName } = this.contextInfo(context);
        const entityName = (entityTypeName ?? "").toLowerCase();

        const lookup = p.contactLookupField.raw?.[0];
        const host: IHostContext = {
            entityName,
            entityDisplayName: this.entityDisplayName || prettifyEntityName(entityName) || "Record",
            recordId: normalizeGuid(entityId),
            contactId: normalizeGuid(lookup?.id),
            contactName: lookup?.name ?? "",
            userPrincipalName: this.userPrincipalName,
        };

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
        } else {
            const key = JSON.stringify(cfg) + (this.userPrincipalName ?? "");
            if (!this.api || key !== this.apiKey) {
                this.apiKey = key;
                this.api = new DocumentApi(cfg.apiBaseUrl, new AuthService(cfg.tenantId, cfg.clientId, cfg.apiScope, this.userPrincipalName));
            }
        }

        const theme = (context.fluentDesignLanguage?.tokenTheme as Theme | undefined) ?? webLightTheme;
        return React.createElement(
            FluentProvider,
            { theme, className: "categorized-doc-manager" },
            React.createElement(DocumentManager, {
                host,
                categories: categoriesFor(entityName, p.categoryConfig.raw),
                api: configError ? null : this.api,
                configError,
            }),
        );
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
        this.api = null;
    }

    private contextInfo(context: ComponentFramework.Context<IInputs>): IContextInfo {
        return ((context.mode as unknown as { contextInfo?: IContextInfo }).contextInfo) ?? {};
    }
}
