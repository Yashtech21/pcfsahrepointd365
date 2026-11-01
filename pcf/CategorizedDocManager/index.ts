import * as React from "react";
import { FluentProvider, webLightTheme, Theme } from "@fluentui/react-components";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { ControlRoot, IControlServices } from "./components/ControlRoot";
import { IApiSettings, SettingsService } from "./services/SettingsService";
import { IHostContext } from "./types";
import { normalizeGuid, prettifyEntityName } from "./utils";

interface IContextInfo { entityId?: string; entityTypeName?: string }

/**
 * Dynamics calls updateView many times while a form loads and on every field change. Everything
 * handed to React here is cached and only replaced when its inputs actually change, so the
 * component tree sees stable props and doesn't refetch or re-render needlessly.
 *
 * Configuration (API connection, categories, per-table settings) is read from Dataverse by
 * ControlRoot via SettingsService; the form properties are optional per-form overrides.
 */
export class CategorizedDocManager implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private context!: ComponentFramework.Context<IInputs>;
    private services!: IControlServices;
    private host: IHostContext | null = null;
    private overrides: Partial<IApiSettings> = {};
    private overridesKey = "\u0000";
    private displayNames = new Map<string, Promise<string>>();

    public init(context: ComponentFramework.Context<IInputs>): void {
        this.context = context;
        this.services = {
            settings: new SettingsService(context.webAPI),
            getEntityDisplayName: this.getEntityDisplayName,
            // The user's UPN is an MSAL login hint so silent SSO picks the right account. Resolved once.
            userPrincipalName: context.webAPI
                .retrieveRecord("systemuser", context.userSettings.userId.replace(/[{}]/g, ""), "?$select=domainname")
                .then(u => (u as { domainname?: string }).domainname || undefined)
                .catch(() => undefined),
        };
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

        const raw = {
            apiBaseUrl: (p.apiBaseUrl.raw ?? "").trim(),
            tenantId: (p.tenantId.raw ?? "").trim(),
            clientId: (p.clientId.raw ?? "").trim(),
            apiScope: (p.apiScope.raw ?? "").trim(),
        };
        const key = JSON.stringify(raw);
        if (key !== this.overridesKey) {
            this.overridesKey = key;
            this.overrides = Object.fromEntries(Object.entries(raw).filter(([, v]) => v)) as Partial<IApiSettings>;
        }

        const boundLookupColumn = ((p.contactLookupField as unknown as { attributes?: { LogicalName?: string } })
            .attributes?.LogicalName ?? "").toLowerCase();

        const theme = (context.fluentDesignLanguage?.tokenTheme as Theme | undefined) ?? webLightTheme;
        return React.createElement(
            FluentProvider,
            { theme, className: "categorized-doc-manager" },
            React.createElement(ControlRoot, {
                host: this.host!,
                services: this.services,
                overrides: this.overrides,
                boundLookupColumn,
            }),
        );
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
        // nothing to release; settings cache is per page load
    }
}
