import { HttpRequest } from "@azure/functions";
import { createRemoteJWKSet, jwtVerify, JWTPayload } from "jose";
import { config } from "./config";

export interface Caller { userId: string; name: string; tenantId: string; roles: string[] }

export class AuthError extends Error {
    constructor(public status: 401 | 403, message: string) { super(message); }
}

let jwks: ReturnType<typeof createRemoteJWKSet> | null = null;

/**
 * Verifies the caller's Entra ID bearer token in code: signature (tenant JWKS), issuer (v1 or v2),
 * audience (this API), expiry, tenant and optional app role.
 *
 * This deliberately does NOT trust X-MS-CLIENT-PRINCIPAL: that header is only safe when App Service
 * Authentication is enabled, and anyone can forge it if it isn't. Easy Auth stays recommended as an
 * extra layer, but security no longer depends on it being configured.
 */
export async function getCaller(req: HttpRequest): Promise<Caller> {
    if (!config.requireAuth) {
        return { userId: "local-dev", name: "local-dev", tenantId: "", roles: [] };
    }
    const header = req.headers.get("authorization") ?? "";
    const match = /^Bearer\s+(.+)$/i.exec(header);
    if (!match) throw new AuthError(401, "Missing bearer token");

    const tenant = config.tenantId.toLowerCase();
    jwks ??= createRemoteJWKSet(new URL(`https://login.microsoftonline.com/${tenant}/discovery/v2.0/keys`), {
        timeoutDuration: 8000,
    });

    let payload: JWTPayload & { tid?: string; oid?: string; name?: string; preferred_username?: string; roles?: string[] };
    try {
        ({ payload } = await jwtVerify(match[1], jwks, {
            issuer: [`https://login.microsoftonline.com/${tenant}/v2.0`, `https://sts.windows.net/${tenant}/`],
            audience: config.apiAudiences,
            clockTolerance: 60,
        }));
    } catch {
        throw new AuthError(401, "Invalid or expired token");
    }

    const caller: Caller = {
        userId: payload.oid ?? payload.sub ?? "",
        name: payload.name ?? payload.preferred_username ?? "",
        tenantId: (payload.tid ?? "").toLowerCase(),
        roles: Array.isArray(payload.roles) ? payload.roles : [],
    };
    if (caller.tenantId !== tenant) throw new AuthError(403, "Caller is not from the expected tenant");
    if (config.requiredRole && !caller.roles.includes(config.requiredRole)) {
        throw new AuthError(403, `Caller lacks the required role ${config.requiredRole}`);
    }
    return caller;
}
