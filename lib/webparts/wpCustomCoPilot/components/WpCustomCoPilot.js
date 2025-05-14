var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
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
import * as React from 'react';
import styles from './WpCustomCoPilot.module.scss';
import { useRef, useEffect } from "react";
import * as ReactWebChat from 'botframework-webchat';
import MSALWrapper from '../SSOAuth/MSALWrapper';
import { Spinner } from '@fluentui/react';
var CoPilotCustomWP = function (props) {
    var botURL = props.botURL, clientID = props.clientID, authority = props.authority, customScope = props.customScope, userDisplayName = props.userDisplayName, botAvatarImage = props.botAvatarImage, botAvatarInitials = props.botAvatarInitials, userEmail = props.userEmail;
    // Check for required properties
    if (!botURL || !clientID || !authority || !customScope) {
        return (React.createElement("section", { className: styles.wpCustomCoPilot },
            React.createElement("div", { style: { textAlign: 'center', padding: '1rem', color: 'red' } }, "Please configure webpart properties")));
    }
    // constructing URL using regional settings
    var environmentEndPoint = botURL.slice(0, botURL.indexOf('/powervirtualagents'));
    var apiVersion = botURL.slice(botURL.indexOf('api-version')).split('=')[1];
    var regionalChannelSettingsURL = "".concat(environmentEndPoint, "/powervirtualagents/regionalchannelsettings?api-version=").concat(apiVersion);
    // Using refs instead of IDs to get the webchat and loading spinner elements
    var webChatRef = useRef(null);
    var loadingSpinnerRef = useRef(null);
    // A utility function that extracts the OAuthCard resource URI from the incoming activity or return undefined
    function getOAuthCardResourceUri(activity) {
        var _a;
        var attachment = (_a = activity === null || activity === void 0 ? void 0 : activity.attachments) === null || _a === void 0 ? void 0 : _a[0];
        if ((attachment === null || attachment === void 0 ? void 0 : attachment.contentType) === 'application/vnd.microsoft.card.oauth' && attachment.content.tokenExchangeResource) {
            return attachment.content.tokenExchangeResource.uri;
        }
    }
    var onDidMount = function () { return __awaiter(void 0, void 0, void 0, function () {
        var MSALWrapperInstance, responseToken, token, regionalChannelURL, regionalResponse, data, directline, response, conversationInfo, store, avatarOptions, canvasStyleOptions;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    MSALWrapperInstance = new MSALWrapper(props.clientID, props.authority);
                    return [4 /*yield*/, MSALWrapperInstance.handleLoggedInUser([props.customScope], props.userEmail)];
                case 1:
                    responseToken = _a.sent();
                    if (!!responseToken) return [3 /*break*/, 3];
                    return [4 /*yield*/, MSALWrapperInstance.acquireAccessToken([props.customScope], props.userEmail)];
                case 2:
                    // Trying to get token if user is not signed-in
                    responseToken = _a.sent();
                    _a.label = 3;
                case 3:
                    token = (responseToken === null || responseToken === void 0 ? void 0 : responseToken.accessToken) || null;
                    return [4 /*yield*/, fetch(regionalChannelSettingsURL)];
                case 4:
                    regionalResponse = _a.sent();
                    if (!regionalResponse.ok) return [3 /*break*/, 6];
                    return [4 /*yield*/, regionalResponse.json()];
                case 5:
                    data = _a.sent();
                    regionalChannelURL = data.channelUrlsById.directline;
                    return [3 /*break*/, 7];
                case 6:
                    console.error("HTTP error! Status: ".concat(regionalResponse.status));
                    _a.label = 7;
                case 7: return [4 /*yield*/, fetch(botURL)];
                case 8:
                    response = _a.sent();
                    if (!response.ok) return [3 /*break*/, 10];
                    return [4 /*yield*/, response.json()];
                case 9:
                    conversationInfo = _a.sent();
                    directline = ReactWebChat.createDirectLine({
                        token: conversationInfo.token,
                        domain: regionalChannelURL + 'v3/directline'
                    });
                    return [3 /*break*/, 11];
                case 10:
                    console.error("HTTP error! Status: ".concat(response.status));
                    _a.label = 11;
                case 11:
                    store = ReactWebChat.createStore({}, function (_a) {
                        var dispatch = _a.dispatch;
                        return function (next) { return function (action) {
                            // Checking whether we should greet the user
                            if (props.greet) {
                                if (action.type === "DIRECT_LINE/CONNECT_FULFILLED") {
                                    console.log("Action:" + action.type);
                                    dispatch({
                                        meta: {
                                            method: "keyboard",
                                        },
                                        payload: {
                                            activity: {
                                                channelData: {
                                                    postBack: true,
                                                },
                                                //Web Chat will show the 'Greeting' System Topic message which has a trigger-phrase 'hello'
                                                name: 'startConversation',
                                                type: "event"
                                            },
                                        },
                                        type: "DIRECT_LINE/POST_ACTIVITY",
                                    });
                                    return next(action);
                                }
                            }
                            // Checking whether the bot is asking for authentication
                            if (action.type === "DIRECT_LINE/INCOMING_ACTIVITY") {
                                var activity = action.payload.activity;
                                if (activity.from && activity.from.role === 'bot' &&
                                    (getOAuthCardResourceUri(activity))) {
                                    directline.postActivity({
                                        type: 'invoke',
                                        name: 'signin/tokenExchange',
                                        value: {
                                            id: activity.attachments[0].content.tokenExchangeResource.id,
                                            connectionName: activity.attachments[0].content.connectionName,
                                            token: token
                                        },
                                        "from": {
                                            id: props.userEmail,
                                            name: props.userFriendlyName,
                                            role: "user"
                                        }
                                    }).subscribe(function (id) {
                                        if (id === "retry") {
                                            // bot was not able to handle the invoke, so display the oauthCard (manual authentication)
                                            console.log("bot was not able to handle the invoke, so display the oauthCard");
                                            return next(action);
                                        }
                                    }, function (error) {
                                        // an error occurred to display the oauthCard (manual authentication)
                                        console.log("An error occurred so display the oauthCard");
                                        return next(action);
                                    });
                                    // token exchange was successful, do not show OAuthCard
                                    return;
                                }
                            }
                            else {
                                return next(action);
                            }
                            return next(action);
                        }; };
                    });
                    avatarOptions = botAvatarImage && botAvatarInitials ? {
                        botAvatarImage: botAvatarImage,
                        botAvatarInitials: botAvatarInitials,
                        userAvatarImage: "/_layouts/15/userphoto.aspx?size=S&username=".concat(userEmail),
                        userAvatarInitials: userDisplayName.charAt(0)
                    } : {};
                    canvasStyleOptions = __assign({ hideUploadButton: true, rootHeight: '100%', rootWidth: '100%', botAvatarBackgroundColor: '#fff', userAvatarBackgroundColor: '#fff', bubbleBackground: '#EBEBED', bubbleTextColor: '#000', bubbleFromUserBackground: '#0057B8', bubbleFromUserTextColor: '#fff', sendBoxBackground: '#F3F4F6' }, avatarOptions);
                    // Render webchat
                    if (token && directline) {
                        if (webChatRef.current && loadingSpinnerRef.current) {
                            webChatRef.current.style.minHeight = '40vh';
                            loadingSpinnerRef.current.style.display = 'none';
                            ReactWebChat.renderWebChat({
                                directLine: directline,
                                store: store,
                                username: userDisplayName,
                                styleOptions: canvasStyleOptions,
                                userID: props.userEmail,
                                sendTypingIndicator: true,
                            }, webChatRef.current);
                        }
                        else {
                            console.error("Webchat or loading spinner not found");
                        }
                    }
                    return [2 /*return*/];
            }
        });
    }); };
    useEffect(function () {
        console.log('Component mounted');
        onDidMount();
        // Cleanup function (optional, similar to componentWillUnmount)
        return function () {
            console.log('Component unmounted');
        };
    }, []); // Empty dependency array ensures this runs only once
    return (React.createElement("section", { className: "".concat(styles.wpCustomCoPilot), style: { width: props.width ? "".concat(props.width, "px") : '100%' } },
        React.createElement("div", { className: styles.header, style: {
                background: props.headerBgColor || '#009FDB',
                color: props.headerTextColor || '#fff',
                padding: '1rem',
                paddingLeft: props.headerPaddingLeft ? "".concat(props.headerPaddingLeft, "px") : '0',
                borderRadius: '4px 4px 0 0',
                fontWeight: 'bold',
                fontSize: props.headerFontSize ? "".concat(props.headerFontSize, "px") : '1.3rem',
                letterSpacing: '0.5px',
                height: props.headerHeight ? "".concat(props.headerHeight, "px") : undefined,
                minHeight: props.headerHeight ? "".concat(props.headerHeight, "px") : undefined,
                maxHeight: props.headerHeight ? "".concat(props.headerHeight, "px") : undefined
            } }, "FiTo AI (Powered by Ask AT&T)"),
        React.createElement("div", { className: styles.chatContainer, id: "chatContainer", style: {
                paddingTop: props.chatContainerPaddingTop ? "".concat(props.chatContainerPaddingTop, "px") : '0',
                height: props.height ? "".concat(props.height, "px") : '400px',
                width: props.width ? "".concat(props.width, "px") : '100%'
            } },
            React.createElement("div", { ref: webChatRef, role: "main", className: styles.webChat }),
            React.createElement("div", { ref: loadingSpinnerRef },
                React.createElement(Spinner, { label: "Loading...", style: { paddingTop: "1rem", paddingBottom: "1rem" } })))));
};
var WpCustomCoPilot = /** @class */ (function (_super) {
    __extends(WpCustomCoPilot, _super);
    function WpCustomCoPilot(props) {
        var _this = _super.call(this, props) || this;
        console.log(props);
        return _this;
    }
    WpCustomCoPilot.prototype.render = function () {
        return (React.createElement(CoPilotCustomWP, __assign({}, this.props)));
    };
    return WpCustomCoPilot;
}(React.Component));
export default WpCustomCoPilot;
//# sourceMappingURL=WpCustomCoPilot.js.map