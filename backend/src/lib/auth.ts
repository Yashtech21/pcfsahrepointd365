import { HttpRequest } from "@azure/functions";
import { config } from "./config";

export interface Caller { userId: string; name: string; tenantId: string; roles: string[] }

interface ClientPrincipal {
    auth_typ?: string;
    claims?: { typ: string; val: string }[];
}

const TENANT_CLAIM = "http://schemas.microsoft.com/identity/claims/tenantid";
const OID_CLAIM = "http://schemas.microsoft.com/identity/claims/objectidentifier";

export class AuthError extends Error {
    constructor(public status: 401 | 403, message: string) { super(message); }
}

/**
 * The Function App is protected by App Service Authentication ("Easy Auth") with Microsoft Entra ID,
 * which validates the bearer token and forwards the caller's claims in X-MS-CLIENT-PRINCIPAL.
 * This check makes sure the request really went through Easy Auth and came from our tenant
 * (and, optionally, that the user holds an app role).
 */
export function getCaller(req: HttpRequest): Caller {
    if (!config.requireEasyAuth) {
        return { userId: "local-dev", name: "local-dev", tenantId: "", roles: [] };
    }
    const header = req.headers.get("x-ms-client-principal");
    if (!header) throw new AuthError(401, "Not authenticated");

    let principal: ClientPrincipal;
    try {
        principal = JSON.parse(Buffer.from(header, "base64").toString("utf8")) as ClientPrincipal;
    } catch {
        throw new AuthError(401, "Invalid client principal");
    }
    const claims = principal.claims ?? [];
    const get = (t: string) => claims.find(c => c.typ === t)?.val ?? "";
    const caller: Caller = {
        userId: get(OID_CLAIM) || get("oid"),
        name: get("name") || get("preferred_username"),
        tenantId: get(TENANT_CLAIM) || get("tid"),
        roles: claims.filter(c => c.typ === "roles" || c.typ.endsWith("/role")).map(c => c.val),
    };

    if (caller.tenantId.toLowerCase() !== config.tenantId.toLowerCase()) {
        throw new AuthError(403, "Caller is not from the expected tenant");
    }
    if (config.requiredRole && !caller.roles.includes(config.requiredRole)) {
        throw new AuthError(403, `Caller lacks the required role ${config.requiredRole}`);
    }
    return caller;
}
