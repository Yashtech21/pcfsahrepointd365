import {
    PublicClientApplication, InteractionRequiredAuthError, AccountInfo, BrowserCacheLocation,
} from "@azure/msal-browser";
import { withTimeout } from "./http";

const SILENT_TIMEOUT_MS = 10000;

/**
 * Acquires an access token for the document API on behalf of the signed-in Dynamics 365 user.
 * The user is already signed in to Entra ID in the browser, so ssoSilent normally succeeds
 * without any prompt; a popup is used only as a fallback (and is not timed, since the user is in it).
 */
export class AuthService {
    private pca: PublicClientApplication;
    private ready: Promise<void>;
    private account: AccountInfo | null = null;

    constructor(tenantId: string, clientId: string, private scope: string, private loginHint: Promise<string | undefined>) {
        this.pca = new PublicClientApplication({
            auth: {
                clientId,
                authority: `https://login.microsoftonline.com/${tenantId}`,
                redirectUri: window.location.origin,
            },
            cache: { cacheLocation: BrowserCacheLocation.SessionStorage },
        });
        this.ready = this.pca.initialize();
    }

    public async getToken(): Promise<string> {
        await withTimeout(this.ready, SILENT_TIMEOUT_MS);
        const scopes = [this.scope];

        this.account ??= this.pca.getAllAccounts()[0] ?? null;
        if (this.account) {
            try {
                return (await withTimeout(this.pca.acquireTokenSilent({ scopes, account: this.account }), SILENT_TIMEOUT_MS)).accessToken;
            } catch (e) {
                if (!(e instanceof InteractionRequiredAuthError)) throw e;
            }
        }
        const loginHint = await this.loginHint;
        try {
            const r = await withTimeout(this.pca.ssoSilent({ scopes, loginHint }), SILENT_TIMEOUT_MS);
            this.account = r.account;
            return r.accessToken;
        } catch {
            const r = await this.pca.acquireTokenPopup({ scopes, loginHint });
            this.account = r.account;
            return r.accessToken;
        }
    }
}
