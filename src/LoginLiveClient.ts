/**
 * LoginLiveClient.ts
 * Copyright: Microsoft 2018
 *
 * Helper client for login.live.com REST API
 */

import * as SyncTasks from 'synctasks';
import { ApiCallOptions, GenericRestClient, WebRequestPriority } from 'simplerestclients';

export interface LoginResponseParams {
    access_token: string;
    expires_in: number;
    refresh_token: string;
    scope: string;
    token_type: string;
    user_id: string;
    foci: string;
}

class LoginLiveClient extends GenericRestClient {
    constructor() {
        super('https://login.live.com/');
    }

    getAccessTokenFromRefreshToken(clientId: string, scope: string, refreshToken: string): SyncTasks.Promise<LoginResponseParams> {
        const params = 'client_id=' + clientId +
            '&scope=' + scope +
            '&refresh_token=' + refreshToken +
            '&grant_type=refresh_token';

        const options: ApiCallOptions = {
            contentType: 'form',
            retries: 5,
            timeout: 300000, // 5 min timeout
            priority: WebRequestPriority.Critical
        };

        return this.performApiPost('oauth20_token.srf', params, options);
    }

    getAccessAndRefreshTokenFromAuthCode(clientId: string, scope: string, redirectUri: string,
            authorizationCode: string): SyncTasks.Promise<LoginResponseParams> {
        const params = 'client_id=' + clientId +
            '&scope=' + scope +
            '&redirect_uri=' + encodeURIComponent(redirectUri) +
            '&code=' + authorizationCode +
            '&grant_type=authorization_code';

        const options: ApiCallOptions = {
            contentType: 'form',
            retries: 5,
            timeout: 300000, // 5 min timeout
            priority: WebRequestPriority.Critical
        };

        return this.performApiPost('oauth20_token.srf', params, options);
    }
}

export default new LoginLiveClient();
