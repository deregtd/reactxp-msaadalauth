/**
* index.native.ts
* Copyright: Microsoft 2018
*
* Native implementation of MSA auth, currently using the OAuthView.
*/

import SyncTasks = require('synctasks');

import { AccessTokenRefreshResult, AppConfig, AuthHelperCommon, MsaNativeRedirectUrl, UnifiedError, UnifiedErrorType,
    UserLoginResult } from './Common';
import LoginLiveClient from './LoginLiveClient';
import MsaOAuthView from './MsaOAuthView';

export class MsaHelper implements AuthHelperCommon {
    constructor(protected _appConfig: AppConfig) {
        // NOOP
    }

    web_getPendingLogin(): UserLoginResult|undefined {
        // WebOnly - NOOP
        return undefined;
    }

    web_getCachedUser(): UserLoginResult|undefined {
        // WebOnly - NOOP
        return undefined;
    }

    web_ackLogin(): void {
        // WebOnly - NOOP
    }

    loginNewUser(scopes: string[], usernameHint?: string): SyncTasks.Promise<UserLoginResult> {
        return MsaOAuthView.login(this._appConfig.clientId, this._appConfig.clientSecret!!!, MsaNativeRedirectUrl,
            scopes, usernameHint)
        .catch(MsaHelper._errCatcher);
    }

    logoutUser(userIdentifier: string, userName: string): SyncTasks.Promise<void> {
        return MsaOAuthView.logout(this._appConfig.clientId, MsaNativeRedirectUrl, userName)
        .catch(MsaHelper._errCatcher);
    }

    getAccessTokenForUserSilent(userIdentifier: string, userName: string, scopes: string[], refreshToken?: string|undefined)
            : SyncTasks.Promise<AccessTokenRefreshResult> {
        if (!refreshToken) {
            return SyncTasks.Rejected({ unifiedType: UnifiedErrorType.Unknown,
                error: 'No refresh token passed to getAccessTokenForUserSilent' } as UnifiedError);
        }

        // TODO: Retries, etc.
        return LoginLiveClient.getAccessTokenFromRefreshToken(this._appConfig.clientId,
                scopes.join(' '), refreshToken).then(parms => {
            const result: AccessTokenRefreshResult = {
                accessToken: {
                    token: parms.access_token,
                    expiresIn: parms.expires_in,
                    scopes: parms.scope.split(' '),
                },
                refreshToken: parms.refresh_token,
            };
            return result;
        }).catch(MsaHelper._errCatcher);
    }

    getAccessTokenForUserInteractive(userIdentifier: string, userName: string, scopes: string[])
            : SyncTasks.Promise<AccessTokenRefreshResult> {
        return MsaOAuthView.login(this._appConfig.clientId, this._appConfig.clientSecret!!!, MsaNativeRedirectUrl,
                scopes, userName).then(loginResult => {
            if (!loginResult.full) {
                return SyncTasks.Rejected<AccessTokenRefreshResult>({ unifiedType: UnifiedErrorType.Unknown,
                    error: 'No full profile in getAccessTokenForUserInteractive' } as UnifiedError);
            }
            if (!loginResult.full.accessToken) {
                return SyncTasks.Rejected<AccessTokenRefreshResult>({ unifiedType: UnifiedErrorType.Unknown,
                    error: 'No accessToken in getAccessTokenForUserInteractive' } as UnifiedError);
            }

            const result: AccessTokenRefreshResult = {
                accessToken: {
                    token: loginResult.full.accessToken.token,
                    expiresIn: loginResult.full.accessToken.expiresIn,
                    scopes: loginResult.full.accessToken.scopes,
                },
                refreshToken: loginResult.full.refreshToken,
            };
            return result;
        }).catch(MsaHelper._errCatcher);
    }

    private static _errCatcher = <T>(error: any) => {
        const errUnified: UnifiedError = {
            unifiedType: MsaHelper._mapErrorType(error),
            error: error,
        };
        return SyncTasks.Rejected<T>(errUnified);
    }

    private static _mapErrorType(error: any): UnifiedErrorType {
        return UnifiedErrorType.Unknown;
    }
}
