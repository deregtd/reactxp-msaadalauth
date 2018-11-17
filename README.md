# ReactXP-MsaAdalAuth

[![GitHub license](https://img.shields.io/badge/license-MIT-blue.svg?style=flat-square)](https://github.com/deregtd/reactxp-msaadalauth/blob/master/LICENSE) [![npm version](https://img.shields.io/npm/v/reactxp-msaadalauth.svg?style=flat-square)](https://www.npmjs.com/package/reactxp-msaadalauth) [![npm downloads](https://img.shields.io/npm/dm/reactxp-msaadalauth.svg?style=flat-square)](https://www.npmjs.com/package/reactxp-msaadalauth) [![David](https://img.shields.io/david/deregtd/reactxp-msaadalauth.svg?style=flat-square)](https://github.com/deregtd/reactxp-msaadalauth)
[![David](https://img.shields.io/david/dev/deregtd/reactxp-msaadalauth.svg?style=flat-square)](https://github.com/deregtd/reactxp-msaadalauth)

> An abstraction around ADAL plus a manual MSA implementation for unified authentication support for ReactXP apps.

## Installation

```shell
npm install --save reactxp-msaadalauth
```

To use with react native platforms, install the appropriate ADAL library for your platform(s):
* https://github.com/AzureAD/azure-activedirectory-library-for-objc
* https://github.com/AzureAD/azure-activedirectory-library-for-android

## Sample Usage

Partial usage example stolen from a private app.

```typescript
class AuthService implements Service {
    private _auth!: MsaAdalAuth;

    startup(): SyncTasks.Promise<void> {
        let adalRedirectUri: string;
        if (AppConfig.getPlatformType() === 'web') {
            adalRedirectUri = window.location.origin;
        } else if (AppConfig.getPlatformType() === 'windows') {
            // Windows uses an auto-generated redirect uri from your app's SID, so this is ignored.
            adalRedirectUri = '';
        } else {
            // Brokered auth
            adalRedirectUri = 'x-msauth-xyz://com.blah.myapp';
        }

        const adalConfig: MsaAdalAppConfig = {
            clientId: (AppConfig.getPlatformType() === 'web') ?
                AdalWebClientId :
                AdalNativeClientId,
            redirectUri: adalRedirectUri,
        };

        const msaConfig: MsaAdalAppConfig = {
            clientId: (AppConfig.getPlatformType() === 'web') ?
                MsaWebClientId :
                MsaNativeClientId,
            clientSecret: (AppConfig.getPlatformType() === 'web') ?
                undefined :
                MsaNativeClientSecret,
            redirectUri: (AppConfig.getPlatformType() === 'web') ?
                window.location.origin :
                '',
        };

        this._auth = new MsaAdalAuth(msaConfig, adalConfig, this._possibleLoginCallback);

        return this._auth.processStartupLogins();
    }

    loginNewUser(): SyncTasks.Promise<void> {
        const msaScopes = [GraphUserUrlScope];
        const adalResourceId = OutlookResourceId;

        // If we have any MSA-logged-in users, start with an AAD login.  On web, MSA will force you to log back out
        // again before letting you do an AAD login.  On native, it auto redirects over to AAD.
        const startWithAdal = _.some(this._users, u => u.enabled && u.isMsa);

        return this._auth.loginNewUser(msaScopes, adalResourceId, !startWithAdal);
    }

    private _possibleLoginCallback = (result: UserLoginResult) => {
        // Partial logins are from MSA, since AAD comes back with full info for us.
        if (result.partial) {
            // See if there's an existing user already for this.
            const existingUser = _.find(this._users, user => user.anchorMailbox === result.partial!!!.anchorMailbox);
            if (existingUser) {
                // See if there's an existing token on the user that matches scopes of the incoming token.
                let newTokenSet = _.clone(existingUser.accessTokens);
                const existingTokenIndex = _.findIndex(existingUser.accessTokens, token =>
                    _.isEqual(token.scopes, result.partial!!!.accessToken.scopes));
                if (existingTokenIndex !== -1) {
                    newTokenSet.splice(existingTokenIndex, 1, result.partial.accessToken);
                } else {
                    newTokenSet.push(result.partial.accessToken);
                }

                const updatedUser = _.extend({}, existingUser, { enabled: true, accessTokens: newTokenSet });
                this._users[updatedUser.userKey] = updatedUser;
                this._trackNewUser(updatedUser, true);
                return this._finishLoginOrUpdate();
            }

            // We have a partial login from MSA, but we need user info to finish it, so query for user info then
            // wrap back around to the login result.
            let promise: SyncTasks.Promise<void>;
            if (AppConfig.getPlatformType() === 'web') {
                promise = GraphClient.getMyInfo(result.partial.accessToken.token).then(info => {
                    const finalResult: UserLoginResult = {
                        full: {
                            userIdentifier: info.id,
                            displayName: info.displayName,
                            email: info.userPrincipalName,
                            anchorMailbox: result.partial!!!.anchorMailbox!!!,
                            refreshToken: result.partial!!!.refreshToken,
                            accessToken: result.partial!!!.accessToken!!!,
                            isMsa: true,
                        }
                    };
                    return this._possibleLoginCallback(finalResult);
                });
            } else {
                promise = ProfileClient.getMyInfo(result.partial.accessToken.token, result.partial.anchorMailbox).then(info => {
                    const finalResult: UserLoginResult = {
                        full: {
                            userIdentifier: result.partial!!!.anchorMailbox,
                            displayName: info.names[0].displayNameDefault,
                            email: info.accounts[0].userPrincipalName,
                            anchorMailbox: result.partial!!!.anchorMailbox,
                            refreshToken: result.partial!!!.refreshToken,
                            accessToken: result.partial!!!.accessToken,
                            isMsa: true,
                        }
                    };
                    return this._possibleLoginCallback(finalResult);
                });
            }
            return promise.catch(() => {
                // TODO: Do we get a username hint with the login so we can at least try again?
                return this.loginNewUser();
            });
        }

        if (result.full) {
            // Full and fresh login to store.
            const user = result.full;

            Instrumentation.log(LogTraceAreas.Auth, 'Logged in user: ' + JSON.stringify(user));

            const userKey = user.userIdentifier as UserKey;

            ServiceStateStore.internal_userLoggedIn(userKey);

            const lastUser = CurrentUsersStore.getUser(userKey);

            const newUser: User = {
                enabled: true,

                userKey,
                accessTokens: _.compact([user.accessToken]),
                refreshToken: user.refreshToken,
                anchorMailbox: user.anchorMailbox,
                fullName: user.displayName,
                email: user.email,
                isMsa: user.isMsa,
            };

            this._trackNewUser(newUser, true);

            return this._finishLoginOrUpdate();
        }

        return SyncTasks.Resolved();
    }

    getAuthToken(user: User, resourceHost: ResourceHost) {
        let applicableScope: string;
        let scopesToFetch: string[];

        if (user.isMsa) {
            applicableScope = GraphUserUrlScope;
            scopesToFetch = [applicableScope];
        } else {
            applicableScope = OutlookResourceId;
            scopesToFetch = [OutlookResourceId];
        }

        const usableToken = _.find(user.accessTokens, token => _.includes(token.scopes, applicableScope));
        if (usableToken) {
            return SyncTasks.Resolved(usableToken.token);
        }

        if (user.isMsa) {
            return this._auth.getMsaAuthToken(user.userKey, user.email, scopesToFetch, user.refreshToken);
        } else {
            return this._auth.getAdalAuthToken(user.userKey, user.email, applicableScope, user.refreshToken);
        }
    }
}
```
