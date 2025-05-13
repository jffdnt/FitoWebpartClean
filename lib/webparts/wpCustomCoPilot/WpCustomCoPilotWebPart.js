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
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import WpCustomCoPilot from './components/WpCustomCoPilot';
var WpCustomCoPilotWebPart = /** @class */ (function (_super) {
    __extends(WpCustomCoPilotWebPart, _super);
    function WpCustomCoPilotWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    WpCustomCoPilotWebPart.prototype.render = function () {
        var element = React.createElement(WpCustomCoPilot, {
            botName: this.properties.botName,
            botURL: this.properties.botURL,
            clientID: this.properties.clientID,
            authority: this.properties.authority,
            customScope: this.properties.customScope,
            greet: this.properties.greet,
            userDisplayName: this.context.pageContext.user.displayName,
            userEmail: this.context.pageContext.user.email,
            userFriendlyName: this.context.pageContext.user.displayName,
            welcomeMessage: this.properties.welcomeMessage,
            botAvatarImage: this.properties.botAvatarImage,
            botAvatarInitials: this.properties.botAvatarInitials
        });
        ReactDom.render(element, this.domElement);
    };
    WpCustomCoPilotWebPart.prototype.onInit = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                return [2 /*return*/, Promise.resolve()];
            });
        });
    };
    WpCustomCoPilotWebPart.prototype.onDispose = function () {
        ReactDom.unmountComponentAtNode(this.domElement);
    };
    Object.defineProperty(WpCustomCoPilotWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    WpCustomCoPilotWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: { description: 'Configure your Copilot Web Part' },
                    groups: [
                        {
                            groupName: 'Bot Settings',
                            groupFields: [
                                PropertyPaneTextField('botName', { label: 'Bot Name' }),
                                PropertyPaneTextField('botURL', { label: 'Bot URL' }),
                                PropertyPaneTextField('clientID', { label: 'Client ID' }),
                                PropertyPaneTextField('authority', { label: 'Authority' }),
                                PropertyPaneTextField('customScope', { label: 'Custom Scope' }),
                                PropertyPaneToggle('greet', { label: 'Greet User' }),
                                PropertyPaneTextField('welcomeMessage', { label: 'Welcome Message' }),
                                PropertyPaneTextField('botAvatarImage', { label: 'Bot Avatar Image' }),
                                PropertyPaneTextField('botAvatarInitials', { label: 'Bot Avatar Initials' })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return WpCustomCoPilotWebPart;
}(BaseClientSideWebPart));
export default WpCustomCoPilotWebPart;
//# sourceMappingURL=WpCustomCoPilotWebPart.js.map