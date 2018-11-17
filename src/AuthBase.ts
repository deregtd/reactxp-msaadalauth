/**
* AuthBase.tsx
* Copyright: Microsoft 2018
*
* Common base class for the web/native abstractions to deal with logging in, getting tokens, etc.
*/

import _ = require('lodash');
import * as SyncTasks from 'synctasks';

import { AccessTokenRefreshResult, AuthHelperCommon, UnifiedError, UnifiedErrorType, UserLoginResult } from './Common';

export abstract class AuthBase {
    protected _adal?: AuthHelperCommon;
    protected _msa?: AuthHelperCommon;

    protected constructor(private _possibleLoginCallback: (result: UserLoginResult) => SyncTasks.Promise<void>|void) {
        // NOOP
    }

    processStartupLogins(): SyncTasks.Promise<void> {
        // This is only activated/useful on ADAL web -- we detect new users that we don't know since the ADAL instance builds
        // them at constructor-time from login redirects...
        const cachedUser = this._adal ? this._adal.web_getCachedUser() : undefined;
        const cachedUserPromise = cachedUser ? this._processLoginResult(cachedUser, [], undefined) : SyncTasks.Resolved();

        return cachedUserPromise.then(() => {
            const pendingLogin = (this._msa && this._msa.web_getPendingLogin()) ||
                (this._adal && this._adal.web_getPendingLogin());
            if (pendingLogin) {
                return this._processLoginResult(pendingLogin, [], undefined);
            }
            return undefined;
        });
    }

    // If you have any MSA-logged-in users, start with an AAD login.  On web, MSA will force you to log back out
    // again before letting you do an AAD login.  On native, it auto redirects over to AAD.
    loginNewUser(msaScopes: string[], adalResourceId: string, startWithMsa = true): SyncTasks.Promise<void> {
        let promise = startWithMsa ? this._msaLogin(msaScopes, adalResourceId) : this._aadLogin(msaScopes, adalResourceId);
        return promise
        .catch((err: UnifiedError) => {
            if (err.unifiedType === UnifiedErrorType.UserCanceled) {
                // User canceled -- swallow it
                return undefined;
            }

            return SyncTasks.Rejected(err);
        });
    }

    private _msaLogin(msaScopes: string[], adalResourceId: string|undefined, usernameHint?: string): SyncTasks.Promise<void> {
        if (!this._msa) {
            return SyncTasks.Rejected('MSA not configured');
        }

        if (!msaScopes.length) {
            return SyncTasks.Rejected('No scopes passed');
        }

        return this._msa.loginNewUser(msaScopes, usernameHint).then(result => {
            return this._processLoginResult(result, msaScopes, adalResourceId);
        });
    }

    private _aadLogin(msaScopes: string[], adalResourceId: string|undefined, usernameHint?: string): SyncTasks.Promise<void> {
        if (!this._adal) {
            return SyncTasks.Rejected('ADAL not configured');
        }

        if (!adalResourceId) {
            return SyncTasks.Rejected('No resource id passed');
        }

        return this._adal.loginNewUser([adalResourceId], usernameHint).then(result => {
            return this._processLoginResult(result, msaScopes, adalResourceId);
        });
    }

    private _processLoginResult(result: UserLoginResult, msaScopes: string[], adalResourceId: string|undefined)
            : SyncTasks.Promise<void> {
        if (result.switchToAadUsername) {
            return this._aadLogin(_.compact([adalResourceId]), result.switchToAadUsername);
        }

        if (result.switchToMsaUsername) {
            return this._msaLogin(msaScopes, result.switchToMsaUsername);
        }

        return SyncTasks.Resolved()
        .then(() => this._possibleLoginCallback(result))
        .then(() => {
            // Wait until the callback promise completes to clean up the weblogins, since there's a common race condition
            // of multiple redirects during login and processing...
            if (this._msa) {
                this._msa.web_ackLogin();
            }
            if (this._adal) {
                this._adal.web_ackLogin();
            }
        });
    }

    getMsaAuthToken(userIdentifier: string, userEmail: string, msaScopes: string[], refreshToken?: string)
            : SyncTasks.Promise<AccessTokenRefreshResult> {
        if (!this._msa) {
            return SyncTasks.Rejected('MSA not configured');
        }

        return this._msa.getAccessTokenForUserSilent(userIdentifier, userEmail, msaScopes, refreshToken)
        .catch((err: UnifiedError) => {
            // See if we need to intercept an interactive_required call.
            if (err.unifiedType === UnifiedErrorType.InteractiveRequired) {
                return this._msa!!!.getAccessTokenForUserInteractive(userIdentifier, userEmail, msaScopes);
            }

            return SyncTasks.Rejected(err);
        });
    }

    getAdalAuthToken(userIdentifier: string, userEmail: string, resourceId: string, refreshToken?: string)
            : SyncTasks.Promise<AccessTokenRefreshResult> {
        if (!this._adal) {
            return SyncTasks.Rejected('ADAL not configured');
        }

        return this._adal.getAccessTokenForUserSilent(userIdentifier, userEmail, [resourceId], refreshToken)
        .catch((err: UnifiedError) => {
            // See if we need to intercept an interactive_required call.
            if (err.unifiedType === UnifiedErrorType.InteractiveRequired) {
                return this._adal!!!.getAccessTokenForUserInteractive(userIdentifier, userEmail, [resourceId]);
            }

            return SyncTasks.Rejected(err);
        });
    }

    logoutMsaUser(userIdentifier: string, userEmail: string) {
        if (!this._msa) {
            return SyncTasks.Rejected('MSA not configured');
        }

        return this._msa.logoutUser(userIdentifier, userEmail);
    }

    logoutAdalUser(userIdentifier: string, userEmail: string) {
        if (!this._adal) {
            return SyncTasks.Rejected('ADAL not configured');
        }

        return this._adal.logoutUser(userIdentifier, userEmail);
    }
}
