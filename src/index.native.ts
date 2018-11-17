/**
* index.native.tsx
* Copyright: Microsoft 2018
*
* Native impl of AuthBase.
*/

import SyncTasks = require('synctasks');

import Adal = require('./adal.native');
import { AuthBase } from './AuthBase';
import { AppConfig, UnifiedError, UnifiedErrorType, UserAccessToken, UserLoginResult } from './Common';
import Msa = require('./msa.native');

export { AppConfig, UnifiedError, UnifiedErrorType, UserAccessToken, UserLoginResult };

export default class AuthNative extends AuthBase {
    constructor(msaConfig: AppConfig|undefined, adalConfig: AppConfig|undefined,
            possibleLoginCallback: (result: UserLoginResult) => SyncTasks.Promise<void>|void) {
        super(possibleLoginCallback);

        if (msaConfig) {
            this._msa = new Msa.MsaHelper(msaConfig);
        }
        if (adalConfig) {
            this._adal = new Adal.AdalHelper(adalConfig);
        }
    }
}
