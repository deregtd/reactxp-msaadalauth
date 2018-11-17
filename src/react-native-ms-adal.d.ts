/**
* react-native-ms-adal.d.ts
* Copyright: Microsoft 2018
*
* Type definitions for the react-native-ms-adal npm package.
*/

declare module 'react-native-ms-adal' {
    interface UserInfo {
        displayableId: string;
        userId: string;
        familyName: string;
        givenName: string;
        identityProvider: string;
        passwordChangeUrl: string;
        passwordExpiresOn?: Date;
        uniqueId: string;
    }

    interface AuthenticationResult {
        accessToken: string;
        accessTokenType: string;
        expiresOn?: Date;
        idToken: string;
        isMultipleResourceRefreshToken: boolean;
        status: string;
        statusCode: number;
        tenantId: string;
        userInfo: UserInfo;

        createAuthorizationHeader(): string;
    }

    interface TokenCacheItem {
        accessToken: string;
        authority: string;
        clientId: string;
        displayableId: string;
        expiresOn?: Date;
        isMultipleResourceRefreshToken: boolean;
        resource: string;
        tenantId: string;
        userInfo: UserInfo;
    }

    interface TokenCache {
        clear(): Promise<void>;
        readItems(): Promise<TokenCacheItem[]>;
        deleteItem(item: TokenCacheItem): Promise<void>;
    }

    export class MSAdalAuthenticationContext {
        authority: string;
        validateAuthority: string;
        tokenCache: TokenCache;

        constructor(authority: string, validateAuthority: boolean);
        static createAsync(authority: string, validateAuthority: boolean): Promise<MSAdalAuthenticationContext>;

        acquireTokenAsync(resourceUrl: string, clientId: string, redirectUrl: string, userIdentifier?: string,
            extraQueryParameters?: string): Promise<AuthenticationResult>;
        acquireTokenSilentAsync(resourceUrl: string, clientId: string, userIdentifier?: string)
            : Promise<AuthenticationResult>;
    }

    export function MSAdalLogin(authority: string, clientId: string, redirectUrl: string, resourceUrl: string)
        : Promise<AuthenticationResult>;
    export function MSAdalLogout(authority: string, redirectUrl: string): Promise<void>;
    export function getValidMSAdalToken(authority: string): AuthenticationResult|undefined;
}
