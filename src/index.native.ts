/**
 * index.native.tsx
 * Copyright: Microsoft 2018
 *
 * Native impl of AuthBase.
 */

import * as SyncTasks from 'synctasks';

import { AuthBase } from './AuthBase';
import { AdalHelper } from './adal.native';
import { AppConfig, UnifiedError, UnifiedErrorType, UserAccessToken, UserLoginResult } from './Common';
import { MsaHelper } from './msa.native';

export { AppConfig, UnifiedError, UnifiedErrorType, UserAccessToken, UserLoginResult };

export default class AuthNative extends AuthBase {
    constructor(msaConfig: AppConfig|undefined, adalConfig: AppConfig|undefined,
            possibleLoginCallback: (result: UserLoginResult) => SyncTasks.Promise<void>|void) {
        super(possibleLoginCallback);

        if (msaConfig) {
            this._msa = new MsaHelper(msaConfig);
        }

        if (adalConfig) {
            this._adal = new AdalHelper(adalConfig);
        }
    }
}
