function req(name: string): string {
    const v = process.env[name];
    if (!v) throw new Error(`Missing required app setting ${name}`);
    return v;
}

export const config = {
    get tenantId() { return req("GRAPH_TENANT_ID"); },
    get clientId() { return req("GRAPH_CLIENT_ID"); },
    get clientSecret() { return req("GRAPH_CLIENT_SECRET"); },
    get spHostname() { return req("SP_HOSTNAME"); },
    get spSitePath() { return req("SP_SITE_PATH"); },
    get spLibraryName() { return process.env.SP_LIBRARY_NAME || "Contacts"; },
    /** Set REQUIRE_AUTH=false only for local development. */
    get requireAuth() { return (process.env.REQUIRE_AUTH ?? "true").toLowerCase() !== "false"; },
    /** Accepted token audiences: the API's App ID URI and its client id (v2 tokens use the GUID). */
    get apiAudiences() {
        const list = (process.env.API_AUDIENCE || "").split(",").map(s => s.trim()).filter(Boolean);
        return list.length ? list : [`api://${req("GRAPH_CLIENT_ID")}`, req("GRAPH_CLIENT_ID")];
    },
    get requiredRole() { return process.env.REQUIRED_ROLE || ""; },
    /** Dataverse org URL, e.g. https://contoso.crm.dynamics.com (source of categories and table configuration). */
    get dataverseUrl() { return req("DATAVERSE_URL"); },
    /** Local development only: read configuration from a seed.json-shaped file instead of Dataverse. */
    get docConfigFile() { return process.env.DOC_CONFIG_FILE || ""; },
    /** How long configuration read from Dataverse is cached, in seconds (default 300). */
    get configCacheSeconds() { return Number(process.env.CONFIG_CACHE_SECONDS || "300"); },
    get maxUploadBytes() { return Number(process.env.MAX_UPLOAD_MB || "100") * 1024 * 1024; },
};
