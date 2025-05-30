var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import { PublicClientApplication, InteractionRequiredAuthError, } from "@azure/msal-browser";
var MSALWrapper = /** @class */ (function () {
    function MSALWrapper(clientId, authority) {
        this.isInitialized = false;
        this.msalConfig = {
            auth: {
                clientId: clientId,
                authority: authority,
            },
            cache: {
                cacheLocation: "localStorage",
            },
        };
        this.msalInstance = new PublicClientApplication(this.msalConfig);
    }
    // Initialize the MSAL instance
    MSALWrapper.prototype.initialize = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!!this.isInitialized) return [3 /*break*/, 2];
                        return [4 /*yield*/, this.msalInstance.initialize()];
                    case 1:
                        _a.sent(); // Ensures initialization
                        this.isInitialized = true;
                        _a.label = 2;
                    case 2: return [2 /*return*/];
                }
            });
        });
    };
    MSALWrapper.prototype.handleLoggedInUser = function (scopes, userEmail) {
        return __awaiter(this, void 0, void 0, function () {
            var userAccount, accounts, accessTokenRequest;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.initialize()];
                    case 1:
                        _a.sent(); // Ensure MSAL is initialized before use
                        userAccount = null;
                        accounts = this.msalInstance.getAllAccounts();
                        if (accounts === null || accounts.length === 0) {
                            console.log("No users are signed in");
                            return [2 /*return*/, null];
                        }
                        else if (accounts.length > 1) {
                            userAccount = this.msalInstance.getAccountByUsername(userEmail);
                        }
                        else {
                            userAccount = accounts[0];
                        }
                        if (userAccount !== null) {
                            accessTokenRequest = {
                                scopes: scopes,
                                account: userAccount,
                            };
                            return [2 /*return*/, this.msalInstance
                                    .acquireTokenSilent(accessTokenRequest)
                                    .then(function (response) {
                                    return response;
                                })
                                    .catch(function (errorinternal) {
                                    console.log(errorinternal);
                                    return null;
                                })];
                        }
                        return [2 /*return*/, null];
                }
            });
        });
    };
    MSALWrapper.prototype.acquireAccessToken = function (scopes, userEmail) {
        return __awaiter(this, void 0, void 0, function () {
            var accessTokenRequest;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.initialize()];
                    case 1:
                        _a.sent(); // Ensure MSAL is initialized before use
                        accessTokenRequest = {
                            scopes: scopes,
                            loginHint: userEmail,
                        };
                        return [2 /*return*/, this.msalInstance
                                .ssoSilent(accessTokenRequest)
                                .then(function (response) {
                                return response;
                            })
                                .catch(function (silentError) {
                                console.log(silentError);
                                if (silentError instanceof InteractionRequiredAuthError) {
                                    return _this.msalInstance
                                        .loginPopup(accessTokenRequest)
                                        .then(function (response) {
                                        return response;
                                    })
                                        .catch(function (error) {
                                        console.log(error);
                                        return null;
                                    });
                                }
                                return null;
                            })];
                }
            });
        });
    };
    return MSALWrapper;
}());
export { MSALWrapper };
export default MSALWrapper;
//# sourceMappingURL=MSALWrapper.js.map