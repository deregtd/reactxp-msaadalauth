/**
 * index.native.ts
 * Copyright: Microsoft 2018
 *
 * React-Native-specific implementation of the adal abstraction.
 */

import * as SyncTasks from 'synctasks';
import compact from 'lodash/compact';
import { MSAdalAuthenticationContext, MSAdalLogout } from 'react-native-ms-adal';

import {
    AccessTokenRefreshResult,
    AppConfig,
    AuthHelperCommon,
    UnifiedError,
    UnifiedErrorType,
    UserLoginResult
} from './Common';

export class AdalHelper implements AuthHelperCommon {
    private _context: MSAdalAuthenticationContext;
    private _redirectUri: string;

    constructor(protected _appConfig: AppConfig) {
        const authority = 'https://login.microsoftonline.com/common';
        this._context = new MSAdalAuthenticationContext(authority, true);
        this._redirectUri = this._appConfig.redirectUri;
    }

    web_getPendingLogin(): UserLoginResult|undefined {
        // Only used on web
        return undefined;
    }

    web_getCachedUser(): UserLoginResult|undefined {
        // Only used on web
        return undefined;
    }

    web_ackLogin(): void {
        // Only used on web
    }

    loginNewUser(scopes: string[], usernameHint?: string): SyncTasks.Promise<UserLoginResult> {
        const extraParms = 'msafed=0' +
            (usernameHint ? '&login_hint=' + encodeURIComponent(usernameHint) : undefined);
        return SyncTasks.fromThenable(this._context.acquireTokenAsync(
            scopes.join(' '), this._appConfig.clientId, this._redirectUri, undefined, extraParms))
        .then(fetchResult => {
            if (fetchResult.userInfo.identityProvider === 'live.com') {
                return {
                    switchToMsaUsername: fetchResult.userInfo.userId,
                } as UserLoginResult;
            }

            const loginResult: UserLoginResult = {
                full: {
                    userIdentifier: fetchResult.userInfo.displayableId || fetchResult.userInfo.uniqueId,
                    displayName: compact([fetchResult.userInfo.givenName, fetchResult.userInfo.familyName]).join(' '),
                    email: fetchResult.userInfo.userId,
                    isMsa: false,
                    anchorMailbox: fetchResult.userInfo.userId,
                    adOid: fetchResult.userInfo.uniqueId,
                    adTid: fetchResult.tenantId,
                    accessToken: {
                        token: fetchResult.accessToken,
                        expiresIn: fetchResult.expiresOn ? fetchResult.expiresOn.getTime() - Date.now() : 0,
                        scopes: scopes,
                    },
                },
            };

            return loginResult;
        })
        .catch(AdalHelper._errCatcher);
    }

    logoutUser(userIdentifier: string, userName: string): SyncTasks.Promise<void> {
        return SyncTasks.fromThenable(MSAdalLogout(this._context.authority, this._redirectUri))
            .catch(AdalHelper._errCatcher);
    }

    getAccessTokenForUserSilent(userIdentifier: string, userName: string, scopes: string[], refreshToken?: string|undefined)
            : SyncTasks.Promise<AccessTokenRefreshResult> {
        return SyncTasks.fromThenable(this._context.acquireTokenSilentAsync(
            scopes.join(' '), this._appConfig.clientId, userIdentifier))
        .then(fetchResult => {
            const tokenResult: AccessTokenRefreshResult = {
                accessToken: {
                    token: fetchResult.accessToken,
                    expiresIn: fetchResult.expiresOn ? fetchResult.expiresOn.getTime() - Date.now() : 0,
                    scopes: scopes,
                }
            };
            return tokenResult;
        })
        .catch(AdalHelper._errCatcher);
    }

    getAccessTokenForUserInteractive(userIdentifier: string, userName: string, scopes: string[])
            : SyncTasks.Promise<AccessTokenRefreshResult> {
        return SyncTasks.fromThenable(this._context.acquireTokenAsync(
            scopes.join(' '), this._appConfig.clientId, this._redirectUri, userIdentifier))
        .then(fetchResult => {
            const tokenResult: AccessTokenRefreshResult = {
                accessToken: {
                    token: fetchResult.accessToken,
                    expiresIn: fetchResult.expiresOn ? fetchResult.expiresOn.getTime() - Date.now() : 0,
                    scopes: scopes,
                }
            };
            return tokenResult;
        })
        .catch(AdalHelper._errCatcher);
    }

    private static _errCatcher = <T>(error: any) => {
        const errUnified: UnifiedError = {
            unifiedType: AdalHelper._mapErrorType(error),
            error: error,
            errorDesc: error && error.message,
        };
        return SyncTasks.Rejected<T>(errUnified);
    }

    private static _mapErrorType(error: any): UnifiedErrorType {
        if (error && error.code) {
            const errorCode = Number(error.code);

            if (errorCode === 0xCAA70004) {
                // UWP "The server or proxy was not found"
                return UnifiedErrorType.ConnectivityIssue;
            }

            if (errorCode === 403) {
                // "The user has cancelled the authorization."
                return UnifiedErrorType.UserCanceled;
            }
            if (errorCode === 200) {
                // "The user credentials are needed to obtain access token. Please call the non-silent acquireTokenWithResource methods."
                return UnifiedErrorType.InteractiveRequired;
            }
            if (errorCode === 211) {
                // "AADSTS50001: Resource 'https://outlook.office.com' is disabled."
                return UnifiedErrorType.CriticalError;
            }
        }

        return UnifiedErrorType.Unknown;
    }
}
