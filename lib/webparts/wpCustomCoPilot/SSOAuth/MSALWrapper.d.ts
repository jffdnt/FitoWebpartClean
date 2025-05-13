import { AuthenticationResult } from "@azure/msal-browser";
export declare class MSALWrapper {
    private msalConfig;
    private msalInstance;
    private isInitialized;
    constructor(clientId: string, authority: string);
    initialize(): Promise<void>;
    handleLoggedInUser(scopes: string[], userEmail: string): Promise<AuthenticationResult | null>;
    acquireAccessToken(scopes: string[], userEmail: string): Promise<AuthenticationResult | null>;
}
export default MSALWrapper;
//# sourceMappingURL=MSALWrapper.d.ts.map