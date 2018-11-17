/**
 * Common.ts
 * Copyright: Microsoft 2018
 *
 * Common interfaces and functionality for msa authentication.
 */

import * as SyncTasks from 'synctasks';

export const MsaNativeRedirectUrl = 'https://login.live.com/oauth20_desktop.srf';
export const MsaAuthorizeUrl = 'https://login.live.com/oauth20_authorize.srf';
export const MsaLogoutUrl = 'https://login.live.com/oauth20_logout.srf';

export interface Dictionary<T> {
    [key: string]: T;
}

export interface UserAccessToken {
    token: string;
    scopes: string[];
    expiresIn: number;
}

interface PartialLoginResult {
    readonly accessToken: UserAccessToken;
    readonly anchorMailbox: string;
    readonly refreshToken?: string;
}

interface FullLoginResult {
    readonly userIdentifier: string;
    readonly displayName: string;
    readonly email: string;
    readonly isMsa: boolean;
    readonly anchorMailbox: string;
    readonly accessToken?: UserAccessToken;
    readonly refreshToken?: string;
}

export interface UserLoginResult {
    readonly switchToAadUsername?: string;
    readonly switchToMsaUsername?: string;
    readonly partial?: PartialLoginResult;
    readonly full?: FullLoginResult;
}

export interface AccessTokenRefreshResult {
    readonly accessToken: UserAccessToken;
    readonly refreshToken?: string;
}

export interface AppConfig {
    clientId: string;
    clientSecret?: string;
    redirectUri: string;
}

export interface AuthHelperCommon {
    // These interfaces only used on web, due to weirdness of the redirect pattern.
    web_getPendingLogin(): UserLoginResult|undefined;
    web_getCachedUser(): UserLoginResult|undefined;
    web_ackLogin(): void;

    loginNewUser(scopes: string[], usernameHint?: string): SyncTasks.Promise<UserLoginResult>;
    logoutUser(userIdentifier: string, userName: string): SyncTasks.Promise<void>;
    getAccessTokenForUserSilent(userIdentifier: string, userName: string, scopes: string[], refreshToken?: string|undefined)
        : SyncTasks.Promise<AccessTokenRefreshResult>;
    getAccessTokenForUserInteractive(userIdentifier: string, userName: string, scopes: string[])
        : SyncTasks.Promise<AccessTokenRefreshResult>;
}

export enum UnifiedErrorType {
    Unknown = -1,

    ConnectivityIssue = 1,
    InteractiveRequired = 2,
    UserCanceled = 3,
    CriticalError = 4,
}

export interface UnifiedError {
    unifiedType: UnifiedErrorType;
    error: any;
    errorDesc?: string;
}
