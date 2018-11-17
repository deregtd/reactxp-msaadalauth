/*
 * MsaOAuthView.tsx
 * Copyright: Microsoft 2018
 *
 * A panel to hold an oauth handler.  Only capable of doing MSA login on web right now.
 */

import * as RX from 'reactxp';
import * as SyncTasks from 'synctasks';
import each from 'lodash/each';
import map from 'lodash/map';

import { Dictionary, MsaAuthorizeUrl, MsaLogoutUrl, UserLoginResult } from './Common';
import LoginLiveClient from './LoginLiveClient';

interface MsaOAuthViewProps extends RX.CommonProps {
    onFinished: (loginResult: UserLoginResult) => void;
    onError: (err: any) => void;

    logout?: boolean;
    clientId: string;
    clientSecret?: string;
    redirectUri: string;
    scopes: string[];
    username?: string;
}

const _styles = {
    container: RX.Styles.createViewStyle({
        flex: 1,
        alignSelf: 'stretch',
        flexDirection: 'column',
        alignItems: 'stretch',
        left: 0,
        right: 0,
        top: 0,
        bottom: 0,
    }),
};

const _modalId = 'MsaOAuthView';

export default class MsaOAuthView extends RX.Component<MsaOAuthViewProps, RX.Stateless> {
    static login(clientId: string, clientSecret: string, redirectUri: string, scopes: string[], usernameHint?: string) {
        let defer = SyncTasks.Defer<UserLoginResult>();
        RX.Modal.show(<MsaOAuthView
            onFinished={ result => defer.resolve(result) }
            onError={ err => defer.reject(err) }
            clientId={ clientId }
            clientSecret={ clientSecret }
            redirectUri={ redirectUri }
            scopes={ scopes }
            username={ usernameHint }
        />, _modalId);
        return defer.promise();
    }

    static logout(clientId: string, redirectUri: string, usernameHint?: string) {
        let defer = SyncTasks.Defer<void>();
        RX.Modal.show(<MsaOAuthView
            onFinished={ result => defer.resolve(undefined) }
            onError={ err => defer.reject(err) }
            clientId={ clientId }
            redirectUri={ redirectUri }
            scopes={ [] }
            logout={ true }
            username={ usernameHint }
        />, _modalId);
        return defer.promise();
    }

    render(): JSX.Element | null {
        return (
            <RX.WebView
                style={ _styles.container }
                url={ this.props.logout ? this._formMSALogoutUrl() : this._formMSALoginUrl() }
                onLoad={ this._onLoad }
                javaScriptEnabled={ true }
                onError={ this._onWebError }
                domStorageEnabled={ true }
                sandbox={ 4095 }
                startInLoadingState={ true }
            />
        );
    }

    private _formMSALoginUrl() {
        let params: Dictionary<string> = {
            'response_type': 'code',
            'scope': ['offline_access', ...this.props.scopes].join(' '),
            'redirect_uri': this.props.redirectUri,
            'client_id': this.props.clientId,
            'aadredir': '1',
        };

        if (this.props.clientSecret) {
            params['client_secret'] = this.props.clientSecret;
        }

        if (this.props.username) {
            params['login_hint'] = this.props.username;
        } else {
            params['prompt'] = 'login';
        }

        return MsaAuthorizeUrl + '?' +
            map(params, (v, k) => k + '=' + encodeURIComponent(v)).join('&');
    }

    private _formMSALogoutUrl() {
        let params: Dictionary<string> = {
            'redirect_uri': this.props.redirectUri,
            'client_id': this.props.clientId,
        };

        if (this.props.username) {
            params['login_hint'] = this.props.username;
        }

        return MsaLogoutUrl + '?' +
            map(params, (v, k) => k + '=' + encodeURIComponent(v)).join('&');
    }

    private _onLoad = (e: RX.Types.SyntheticEvent) => {
        if (this.props.logout) {
            // Any page load means the logout is finished.
            this.props.onFinished({});
            this._dismiss();
            return;
        }

        const url = e.nativeEvent.url as string;
        if (url.substr(0, this.props.redirectUri.length) === this.props.redirectUri) {
            let parsedParts: Dictionary<string> = {};
            each(url.substr(this.props.redirectUri.length + 1).split('&'), p => {
                const bits = p.split('=');
                parsedParts[bits[0]] = bits[1] ? decodeURIComponent(bits[1]) : '';
            });

            if (parsedParts['error']) {
                if (parsedParts['error'] === 'aad_auth') {
                    this._dismiss();
                    this.props.onFinished({ switchToAadUsername: parsedParts['username'] });
                    return;
                }
                if (parsedParts['error'] === 'access_denied') {
                    // User canceled
                    this._dismiss();
                    this.props.onFinished({});
                    return;
                }

                this._dismiss();
                this.props.onError({ error: parsedParts['error'], errorDesc: parsedParts['error_description'] });
                return;
            }

            // TODO: Validate State
            const code = parsedParts['code'];
            LoginLiveClient.getAccessAndRefreshTokenFromAuthCode(this.props.clientId, this.props.scopes.join(' '),
                this.props.redirectUri, code)
            .then(params => {
                const anchorMailbox = 'CID:' + params.user_id;
                const user: UserLoginResult = {
                    partial: {
                        anchorMailbox,
                        accessToken: {
                            token: params.access_token,
                            scopes: params.scope.split(' '),
                            expiresIn: params.expires_in,
                        },
                        refreshToken: params.refresh_token,
                    }
                };
                this._dismiss();
                this.props.onFinished(user);
            }).catch(err => {
                this._dismiss();
                this.props.onError(err);
            });
        }
    }

    private _onWebError = (err: any) => {
        this._dismiss();
        this.props.onError(err);
    }

    private _dismiss() {
        RX.Modal.dismiss(_modalId);
    }
}
