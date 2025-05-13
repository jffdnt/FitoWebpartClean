import { WebPartContext } from '@microsoft/sp-webpart-base';
export interface ICoPilotConfiguration {
    botURL: string;
    botName?: string;
    buttonLabel?: string;
    botAvatarImage?: string;
    botAvatarInitials?: string;
    greet?: boolean;
    customScope: string;
    clientID: string;
    authority: string;
    cacheTimeout?: number;
    errorRetryAttempts?: number;
}
export declare class ConfigurationService {
    private context;
    private static CONFIG_LIST_NAME;
    private static CACHE_KEY;
    private static DEFAULT_CACHE_TIMEOUT;
    private static DEFAULT_RETRY_ATTEMPTS;
    private static DEFAULT_RETRY_DELAY;
    private cachedConfig;
    private cacheTimestamp;
    private graphClient;
    constructor(context: WebPartContext);
    private ensureGraphClient;
    getConfiguration(): Promise<ICoPilotConfiguration>;
    private getConfigFromListWithRetry;
    private getConfigFromList;
    private checkIfListExists;
    private getListId;
    private mapListItemToConfig;
    private initializeCache;
    private getFromCache;
    private updateCache;
    private getFallbackConfiguration;
    private delay;
}
//# sourceMappingURL=ConfigurationService.d.ts.map