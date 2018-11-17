/**
 * index.web.ts
 * Copyright: Microsoft 2018
 *
 * web-specific implementation of the hellojs abstraction.
 */

import * as SyncTasks from 'synctasks';
import each from 'lodash/each';
import extend from 'lodash/extend';
import map from 'lodash/map';

import {
    AccessTokenRefreshResult,
    AppConfig,
    AuthHelperCommon,
    Dictionary,
    MsaAuthorizeUrl,
    MsaLogoutUrl,
    UserLoginResult
} from './Common';

const LocalStoragePendingLoginKey = 'msaLoginPending';

export class MsaHelper implements AuthHelperCommon {
    private _pendingLogin: UserLoginResult|undefined;

    constructor(private _appConfig: AppConfig) {
        const pending = window.localStorage.getItem(LocalStoragePendingLoginKey);

        // Don't immediately init unless we were pending a login, since it might eat the AAD login attempt.
        if (pending === '1') {
            if (window.location.hash) {
                let parsedParts: Dictionary<string> = {};
                each(window.location.hash.substr(1).split('&'), p => {
                    const bits = p.split('=');
                    parsedParts[bits[0]] = bits[1] ? decodeURIComponent(bits[1]) : '';
                });

                if (parsedParts['error']) {
                    window.localStorage.setItem(LocalStoragePendingLoginKey, '');
                    if (parsedParts['error'] === 'aad_auth') {
                        this._pendingLogin = {
                            switchToAadUsername: parsedParts['username']
                        };
                    } else {
                        alert('MSA Login Error: ' + parsedParts['error'] + '\nDescription: ' + parsedParts['error_description']);
                    }
                } else if (parsedParts['access_token']) {
                    this._pendingLogin = {
                        partial: {
                            anchorMailbox: 'CID:' + parsedParts['user_id'],
                            accessToken: {
                                token: parsedParts['access_token'],
                                expiresIn: Number(parsedParts['expires_in']),
                                scopes: parsedParts['scope'].split(' '),
                            }
                        }
                    };
                }
            } else {
                // Nothing in the URL, so clear the pending flag.
                window.localStorage.setItem(LocalStoragePendingLoginKey, '');
            }
        }
    }

    web_getPendingLogin(): UserLoginResult|undefined {
        return this._pendingLogin;
    }

    web_getCachedUser(): UserLoginResult|undefined {
        // Only used by ADAL
        return undefined;
    }

    // Need a separate ack because it may load several times due to wacky oauth redirects before we finally
    // process the login and store it.
    web_ackLogin(): void {
        window.localStorage.setItem(LocalStoragePendingLoginKey, '');
    }

    loginNewUser(scopes: string[], usernameHint?: string): SyncTasks.Promise<UserLoginResult> {
        window.localStorage.setItem(LocalStoragePendingLoginKey, '1');

        let extraParams: Dictionary<string> = { prompt: 'login' };
        if (usernameHint) {
            extraParams['login_hint'] = usernameHint;
        }
        window.location.href = this._formMSALoginUrl(scopes, extraParams);

        return SyncTasks.Defer().promise();
    }

    private _formMSALoginUrl(scopes: string[], extraParams?: Dictionary<string>) {
        let params: Dictionary<string> = extend({
            'response_type': 'token',
            'scope': scopes.join(' '),
            'redirect_uri': this._appConfig.redirectUri,
            'client_id': this._appConfig.clientId,
            'aadredir': '1',
        }, extraParams);

        return MsaAuthorizeUrl + '?' +
            map(params, (v, k) => k + '=' + encodeURIComponent(v)).join('&');
    }

    private _formMSALogoutUrl(loginHint?: string) {
        let params: Dictionary<string> = {
            'redirect_uri': window.location.origin,
            'client_id': this._appConfig.clientId,
        };

        if (loginHint) {
            params['login_hint'] = loginHint;
        }

        return MsaLogoutUrl + '?' +
            map(params, (v, k) => k + '=' + encodeURIComponent(v)).join('&');
    }

    logoutUser(userIdentifier: string, userName: string): SyncTasks.Promise<void> {
        window.location.href = this._formMSALogoutUrl(userName);

        return SyncTasks.Defer<void>().promise();
    }

    getAccessTokenForUserSilent(userIdentifier: string, userName: string, scopes: string[],
            refreshToken?: string|undefined): SyncTasks.Promise<AccessTokenRefreshResult> {

        // TODO: Support silent refresh
        return this.getAccessTokenForUserInteractive(userIdentifier, userName, scopes);
    }

    getAccessTokenForUserInteractive(userIdentifier: string, userName: string, scopes: string[]) {
        window.localStorage.setItem(LocalStoragePendingLoginKey, '1');

        window.location.href = this._formMSALoginUrl(scopes, { login_hint: userName });

        return SyncTasks.Defer<AccessTokenRefreshResult>().promise();
    }

    // private static _errCatcher = <T>(error: any) => {
    //     const errUnified: UnifiedError = {
    //         unifiedType: MsaHelper._mapErrorType(error),
    //         error,
    //     };

    //     return SyncTasks.Rejected<T>(errUnified);
    // }

    // private static _mapErrorType(error: any): UnifiedErrorType {
    //     return UnifiedErrorType.Unknown;
    // }
}
