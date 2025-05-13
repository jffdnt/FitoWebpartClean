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
import { Log } from '@microsoft/sp-core-library';
var LOG_SOURCE = 'ConfigurationService';
var ConfigurationService = /** @class */ (function () {
    function ConfigurationService(context) {
        this.cachedConfig = null;
        this.cacheTimestamp = 0;
        this.context = context;
        this.initializeCache();
    }
    ConfigurationService.prototype.ensureGraphClient = function () {
        return __awaiter(this, void 0, void 0, function () {
            var _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (!!this.graphClient) return [3 /*break*/, 2];
                        _a = this;
                        return [4 /*yield*/, this.context.msGraphClientFactory.getClient('3')];
                    case 1:
                        _a.graphClient = _b.sent();
                        _b.label = 2;
                    case 2: return [2 /*return*/];
                }
            });
        });
    };
    ConfigurationService.prototype.getConfiguration = function () {
        return __awaiter(this, void 0, void 0, function () {
            var cachedConfig, listConfig, error_1, fallbackConfig;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        cachedConfig = this.getFromCache();
                        if (cachedConfig) {
                            Log.info(LOG_SOURCE, 'Retrieved configuration from cache');
                            return [2 /*return*/, cachedConfig];
                        }
                        return [4 /*yield*/, this.ensureGraphClient()];
                    case 1:
                        _a.sent();
                        return [4 /*yield*/, this.getConfigFromListWithRetry()];
                    case 2:
                        listConfig = _a.sent();
                        if (listConfig) {
                            this.updateCache(listConfig);
                            return [2 /*return*/, listConfig];
                        }
                        throw new Error('No valid configuration found');
                    case 3:
                        error_1 = _a.sent();
                        Log.error(LOG_SOURCE, error_1);
                        fallbackConfig = this.getFallbackConfiguration();
                        if (fallbackConfig) {
                            return [2 /*return*/, fallbackConfig];
                        }
                        throw error_1;
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    ConfigurationService.prototype.getConfigFromListWithRetry = function () {
        var _a;
        return __awaiter(this, void 0, void 0, function () {
            var retryAttempts, attempt, config, error_2;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        retryAttempts = ((_a = this.cachedConfig) === null || _a === void 0 ? void 0 : _a.errorRetryAttempts) || ConfigurationService.DEFAULT_RETRY_ATTEMPTS;
                        attempt = 0;
                        _b.label = 1;
                    case 1:
                        if (!(attempt < retryAttempts)) return [3 /*break*/, 7];
                        _b.label = 2;
                    case 2:
                        _b.trys.push([2, 4, , 6]);
                        return [4 /*yield*/, this.getConfigFromList()];
                    case 3:
                        config = _b.sent();
                        if (config)
                            return [2 /*return*/, config];
                        return [3 /*break*/, 7];
                    case 4:
                        error_2 = _b.sent();
                        if (attempt === retryAttempts - 1)
                            throw error_2;
                        return [4 /*yield*/, this.delay(ConfigurationService.DEFAULT_RETRY_DELAY * Math.pow(2, attempt))];
                    case 5:
                        _b.sent();
                        return [3 /*break*/, 6];
                    case 6:
                        attempt++;
                        return [3 /*break*/, 1];
                    case 7: return [2 /*return*/, null];
                }
            });
        });
    };
    ConfigurationService.prototype.getConfigFromList = function () {
        return __awaiter(this, void 0, void 0, function () {
            var listExists, siteId, listId, response, error_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 4, , 5]);
                        return [4 /*yield*/, this.checkIfListExists()];
                    case 1:
                        listExists = _a.sent();
                        if (!listExists) {
                            Log.info(LOG_SOURCE, 'Configuration list does not exist');
                            return [2 /*return*/, null];
                        }
                        siteId = this.context.pageContext.site.id.toString();
                        return [4 /*yield*/, this.getListId()];
                    case 2:
                        listId = _a.sent();
                        if (!listId)
                            return [2 /*return*/, null];
                        return [4 /*yield*/, this.graphClient.api("/sites/".concat(siteId, "/lists/").concat(listId, "/items"))
                                .expand('fields')
                                .orderby('fields/Modified desc')
                                .top(1)
                                .get()];
                    case 3:
                        response = _a.sent();
                        if (!response.value || response.value.length === 0) {
                            return [2 /*return*/, null];
                        }
                        return [2 /*return*/, this.mapListItemToConfig(response.value[0].fields)];
                    case 4:
                        error_3 = _a.sent();
                        Log.error(LOG_SOURCE, new Error("Error getting configuration from list: ".concat(error_3)));
                        return [2 /*return*/, null];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    ConfigurationService.prototype.checkIfListExists = function () {
        return __awaiter(this, void 0, void 0, function () {
            var siteId, response, _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 2, , 3]);
                        siteId = this.context.pageContext.site.id.toString();
                        return [4 /*yield*/, this.graphClient.api("/sites/".concat(siteId, "/lists"))
                                .filter("displayName eq '".concat(ConfigurationService.CONFIG_LIST_NAME, "'"))
                                .get()];
                    case 1:
                        response = _b.sent();
                        return [2 /*return*/, response.value && response.value.length > 0];
                    case 2:
                        _a = _b.sent();
                        return [2 /*return*/, false];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    ConfigurationService.prototype.getListId = function () {
        return __awaiter(this, void 0, void 0, function () {
            var siteId, response, error_4;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        siteId = this.context.pageContext.site.id.toString();
                        return [4 /*yield*/, this.graphClient.api("/sites/".concat(siteId, "/lists"))
                                .filter("displayName eq '".concat(ConfigurationService.CONFIG_LIST_NAME, "'"))
                                .select('id')
                                .get()];
                    case 1:
                        response = _a.sent();
                        if (response.value && response.value.length > 0) {
                            return [2 /*return*/, response.value[0].id];
                        }
                        return [2 /*return*/, null];
                    case 2:
                        error_4 = _a.sent();
                        Log.error(LOG_SOURCE, new Error("Error getting list ID: ".concat(error_4)));
                        return [2 /*return*/, null];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    ConfigurationService.prototype.mapListItemToConfig = function (fields) {
        return {
            botURL: fields.BotURL,
            botName: fields.Title,
            buttonLabel: fields.ButtonLabel,
            botAvatarImage: fields.BotAvatarImage,
            botAvatarInitials: fields.BotAvatarInitials,
            greet: fields.Greet,
            customScope: fields.CustomScope,
            clientID: fields.ClientID,
            authority: fields.Authority,
            cacheTimeout: fields.CacheTimeout,
            errorRetryAttempts: fields.ErrorRetryAttempts
        };
    };
    ConfigurationService.prototype.initializeCache = function () {
        try {
            var cached = localStorage.getItem(ConfigurationService.CACHE_KEY);
            if (cached) {
                var _a = JSON.parse(cached), config = _a.config, timestamp = _a.timestamp;
                this.cachedConfig = config;
                this.cacheTimestamp = timestamp;
            }
        }
        catch (_b) {
            Log.warn(LOG_SOURCE, 'Failed to initialize cache from localStorage');
        }
    };
    ConfigurationService.prototype.getFromCache = function () {
        if (!this.cachedConfig)
            return null;
        var now = Date.now();
        var cacheTimeout = (this.cachedConfig.cacheTimeout || ConfigurationService.DEFAULT_CACHE_TIMEOUT) * 60 * 1000;
        if (now - this.cacheTimestamp > cacheTimeout) {
            Log.info(LOG_SOURCE, 'Cache expired');
            return null;
        }
        return this.cachedConfig;
    };
    ConfigurationService.prototype.updateCache = function (config) {
        this.cachedConfig = config;
        this.cacheTimestamp = Date.now();
        try {
            localStorage.setItem(ConfigurationService.CACHE_KEY, JSON.stringify({
                config: config,
                timestamp: this.cacheTimestamp
            }));
        }
        catch (_a) {
            Log.warn(LOG_SOURCE, 'Failed to update localStorage cache');
        }
    };
    ConfigurationService.prototype.getFallbackConfiguration = function () {
        try {
            var fallback = localStorage.getItem(ConfigurationService.CACHE_KEY);
            if (fallback) {
                var config = JSON.parse(fallback).config;
                Log.info(LOG_SOURCE, 'Using fallback configuration from localStorage');
                return config;
            }
        }
        catch (_a) {
            Log.error(LOG_SOURCE, new Error('Failed to get fallback configuration'));
        }
        return null;
    };
    ConfigurationService.prototype.delay = function (ms) {
        return new Promise(function (resolve) { return setTimeout(resolve, ms); });
    };
    ConfigurationService.CONFIG_LIST_NAME = 'CopilotAgentConfig';
    ConfigurationService.CACHE_KEY = 'CHATBOT_CONFIG_CACHE';
    ConfigurationService.DEFAULT_CACHE_TIMEOUT = 30;
    ConfigurationService.DEFAULT_RETRY_ATTEMPTS = 3;
    ConfigurationService.DEFAULT_RETRY_DELAY = 1000;
    return ConfigurationService;
}());
export { ConfigurationService };
//# sourceMappingURL=ConfigurationService.js.map