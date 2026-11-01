import {
    PublicClientApplication, InteractionRequiredAuthError, AccountInfo, BrowserCacheLocation,
} from "@azure/msal-browser";

/**
 * Acquires an access token for the document API on behalf of the signed-in Dynamics 365 user.
 * The user is already signed in to Entra ID in the browser, so ssoSilent normally succeeds
 * without any prompt; a popup is used only as a fallback.
 */
export class AuthService {
    private pca: PublicClientApplication;
    private ready: Promise<void>;
    private account: AccountInfo | null = null;

    constructor(tenantId: string, clientId: string, private scope: string, private loginHint?: string) {
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
        await this.ready;
        const scopes = [this.scope];

        if (!this.account) {
            this.account = this.pca.getAllAccounts()[0] ?? null;
        }
        if (this.account) {
            try {
                return (await this.pca.acquireTokenSilent({ scopes, account: this.account })).accessToken;
            } catch (e) {
                if (!(e instanceof InteractionRequiredAuthError)) throw e;
            }
        }
        try {
            const r = await this.pca.ssoSilent({ scopes, loginHint: this.loginHint });
            this.account = r.account;
            return r.accessToken;
        } catch {
            const r = await this.pca.acquireTokenPopup({ scopes, loginHint: this.loginHint });
            this.account = r.account;
            return r.accessToken;
        }
    }
}
