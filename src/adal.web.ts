/**
* index.web.ts
* Copyright: Microsoft 2018
*
* web-specific implementation of the msal abstraction.
*/

import AdalAuthenticationContext = require('adal-angular');
import _ = require('lodash');
import SyncTasks = require('synctasks');

import { AccessTokenRefreshResult, AppConfig, AuthHelperCommon, UnifiedError, UnifiedErrorType,
    UserLoginResult } from './Common';

interface AdalWebProfile {
    aio?: string;
    altsecid?: string;
    amr?: string[];
    aud?: string;
    email: string;
    exp: number;
    family_name?: string;
    given_name?: string;
    iat: number;
    ipaddr: string;
    iss: string;
    nonce?: string;
    sub?: string;
    tid?: string;
    unique_name: string;
    uti: string;
    ver: string;
}

interface AdalWebUserInfo {
    userName: string;
    profile: AdalWebProfile;
}

const LocalStoragePendingLoginKey = 'adalLoginPending';

export class AdalHelper implements AuthHelperCommon {
    private _pendingLogin: UserLoginResult|undefined;

    private _context: AdalAuthenticationContext|undefined;

    constructor(private _appConfig: AppConfig) {
        const pending = window.localStorage.getItem(LocalStoragePendingLoginKey);

        // Don't immediately init unless we were pending a login, since it might eat the MSA login attempt.
        if (pending === '1') {
            if (window.location.search) {
                const paramList = window.location.search.substr(1).split('&');
                let params: _.Dictionary<string> = {};
                _.each(paramList, p => {
                    const x = p.split('=');
                    params[x[0]] = decodeURIComponent(x[1]);
                });
                if (params['error'] === 'msa_auth') {
                    // Clear pending ADAL login so it won't start up and eat the url.
                    window.localStorage.setItem(LocalStoragePendingLoginKey, '');
                    // Clear the ?error from the url so we don't restore it after logging in with MSA
                    window.history.pushState(null, '', window.location.origin);
                    // Redirect away to MSA
                    this._pendingLogin = {
                        switchToMsaUsername: params['username']
                    };
                }
            } else if (!window.location.hash) {
                // Nothing in the URL, so clear the pending flag.
                window.localStorage.setItem(LocalStoragePendingLoginKey, '');
            }

            this._initContext();
        }

        if (window !== window.parent) {
            // in an iframe, bail
            this._initContext();
            location.href = 'about:blank';
            return;
        }
    }

    private _initContext() {
        if (this._context) {
            return;
        }

        const options: AdalAuthenticationContext.Options = {
            clientId: this._appConfig.clientId,
            cacheLocation: 'localStorage',
            popUp: false,
            redirectUri: window.location.origin,
            navigateToLoginRequestUrl: false,
        };
        this._context = new AdalAuthenticationContext(options);
        this._context.handleWindowCallback();
    }

    private _mapUserInfo(info: AdalWebUserInfo): UserLoginResult {
        const userInfo: UserLoginResult = {
            full: {
                userIdentifier: info.profile.unique_name,
                displayName: _.compact([info.profile.given_name, info.profile.family_name]).join(' '),
                email: info.userName,
                anchorMailbox: info.userName,
                isMsa: false,
            }
        };

        return userInfo;
    }

    web_getPendingLogin(): UserLoginResult|undefined {
        return this._pendingLogin;
    }

    web_getCachedUser(): UserLoginResult|undefined {
        if (!this._context) {
            return undefined;
        }

        const cachedUser = this._context.getCachedUser();
        if (!cachedUser) {
            return undefined;
        }

        return this._mapUserInfo(cachedUser);
    }

    // Need a separate ack because it may load several times due to wacky oauth redirects before we finally
    // process the login and store it.
    web_ackLogin(): void {
        window.localStorage.setItem(LocalStoragePendingLoginKey, '');
    }

    loginNewUser(scopes: string[], usernameHint?: string): SyncTasks.Promise<UserLoginResult> {
        this._initContext();

        window.localStorage.setItem(LocalStoragePendingLoginKey, '1');

        let eqp = 'msaredir=1&msafed=0';
        if (usernameHint) {
             eqp += '&login_hint=' + usernameHint;
        }
        this._context!!!.config.extraQueryParameter = eqp;

        // TODO: Verify this does something.
        this._context!!!.config.loginResource = scopes.join(' ');

        this._context!!!.login();

        this._context!!!.config.extraQueryParameter = '';

        return SyncTasks.Defer<UserLoginResult>().promise();
    }

    logoutUser(userIdentifier: string, userName: string): SyncTasks.Promise<void> {
        this._initContext();

        this._context!!!.clearCache();
        this._context!!!.logOut();

        return SyncTasks.Defer<void>().promise();
    }

    getAccessTokenForUserSilent(userIdentifier: string, userName: string, scopes: string[], refreshToken?: string|undefined)
            : SyncTasks.Promise<AccessTokenRefreshResult> {
        let defer = SyncTasks.Defer<AccessTokenRefreshResult>();

        this._initContext();

        this._context!!!.config.extraQueryParameter = 'login_hint=' + userName;

        this._context!!!.acquireToken(scopes.join(' '), (errorDesc: string | null, token: string | null, error: any) => {
            if (errorDesc || error) {
                defer.reject(AdalHelper._errCatcher(error, errorDesc));
                return;
            }

            defer.resolve({ accessToken: { token: token!!!, scopes: scopes, expiresIn: 0 } });
        });

        this._context!!!.config.extraQueryParameter = '';

        return defer.promise();
    }

    getAccessTokenForUserInteractive(userIdentifier: string, userName: string, scopes: string[])
            : SyncTasks.Promise<AccessTokenRefreshResult> {
        let defer = SyncTasks.Defer<AccessTokenRefreshResult>();

        this._initContext();

        this._context!!!.acquireTokenPopup(scopes.join(' '), null, null,
                (errorDesc: string | null, token: string | null, error: any) => {
            if (errorDesc || error) {
                this._context!!!.acquireTokenRedirect(scopes.join(' '));

                // Hang it.
                return;
            }

            defer.resolve({ accessToken: { token: token!!!, scopes: scopes, expiresIn: 0 } });
        });

        return defer.promise();
    }

    private static _errCatcher = <T>(error: any, errorDesc: string|null) => {
        const errUnified: UnifiedError = {
            unifiedType: AdalHelper._mapErrorType(error, errorDesc),
            error,
            errorDesc: errorDesc || undefined,
        };

        return errUnified;
    }

    private static _mapErrorType(error: any, errorDesc: string|null): UnifiedErrorType {
        if (error === 'interaction_required' || error === 'login required' || error === 'login_required') {
            return UnifiedErrorType.InteractiveRequired;
        }

        if (error === 'Token Renewal Failed' && errorDesc === 'Token renewal operation failed due to timeout') {
            return UnifiedErrorType.ConnectivityIssue;
        }

        return UnifiedErrorType.Unknown;
    }
}
