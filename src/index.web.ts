/**
 * index.web.tsx
 * Copyright: Microsoft 2018
 *
 * Web impl of AuthBase.
 */

import * as SyncTasks from 'synctasks';

import { AdalHelper } from './adal.web';
import { AuthBase } from './AuthBase';
import { AppConfig, UnifiedError, UnifiedErrorType, UserAccessToken, UserLoginResult } from './Common';
import { MsaHelper } from './msa.web';

export { AppConfig, UnifiedError, UnifiedErrorType, UserAccessToken, UserLoginResult };

export default class AuthWeb extends AuthBase {
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
