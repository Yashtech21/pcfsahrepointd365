export interface IDocumentItem {
    id: string;
    name: string;
    category: string;
    sourceEntity: string;
    sourceRecordId: string;
    uploadDate: string;
    webUrl: string;
    size: number;
    fileType: string;
}

export interface IControlConfig {
    apiBaseUrl: string;
    tenantId: string;
    clientId: string;
    apiScope: string;
}

export interface IHostContext {
    entityName: string;          // logical name of the hosting entity, e.g. "opportunity"
    entityDisplayName: string;   // e.g. "Opportunity"
    recordId: string;            // lowercase GUID of the hosting record ("" on create form)
    contactId: string;           // lowercase GUID of the bound contact ("" when empty)
    contactName: string;
    userPrincipalName?: string;
}
