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
    get requireEasyAuth() { return (process.env.REQUIRE_EASY_AUTH ?? "true").toLowerCase() !== "false"; },
    get requiredRole() { return process.env.REQUIRED_ROLE || ""; },
    get maxUploadBytes() { return Number(process.env.MAX_UPLOAD_MB || "100") * 1024 * 1024; },
};
