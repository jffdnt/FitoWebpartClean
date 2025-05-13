"use strict";

var _interopRequireDefault = require("@babel/runtime/helpers/interopRequireDefault");

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.default = void 0;

var _defineProperty2 = _interopRequireDefault(require("@babel/runtime/helpers/defineProperty"));

var _botframeworkWebchatApi = require("botframework-webchat-api");

var _defaultStyleOptions = _interopRequireDefault(require("./adaptiveCards/defaultStyleOptions"));

function ownKeys(object, enumerableOnly) { var keys = Object.keys(object); if (Object.getOwnPropertySymbols) { var symbols = Object.getOwnPropertySymbols(object); enumerableOnly && (symbols = symbols.filter(function (sym) { return Object.getOwnPropertyDescriptor(object, sym).enumerable; })), keys.push.apply(keys, symbols); } return keys; }

function _objectSpread(target) { for (var i = 1; i < arguments.length; i++) { var source = null != arguments[i] ? arguments[i] : {}; i % 2 ? ownKeys(Object(source), !0).forEach(function (key) { (0, _defineProperty2.default)(target, key, source[key]); }) : Object.getOwnPropertyDescriptors ? Object.defineProperties(target, Object.getOwnPropertyDescriptors(source)) : ownKeys(Object(source)).forEach(function (key) { Object.defineProperty(target, key, Object.getOwnPropertyDescriptor(source, key)); }); } return target; }

var fullBundleDefaultStyleOptions = _objectSpread(_objectSpread({}, _botframeworkWebchatApi.defaultStyleOptions), _defaultStyleOptions.default);

var _default = fullBundleDefaultStyleOptions;
exports.default = _default;
//# sourceMappingURL=data:application/json;charset=utf-8;base64,eyJ2ZXJzaW9uIjozLCJuYW1lcyI6WyJmdWxsQnVuZGxlRGVmYXVsdFN0eWxlT3B0aW9ucyIsImRlZmF1bHRTdHlsZU9wdGlvbnMiLCJhZGFwdGl2ZUNhcmRzRGVmYXVsdFN0eWxlT3B0aW9ucyJdLCJzb3VyY2VSb290IjoiYnVuZGxlOi8vLyIsInNvdXJjZXMiOlsiLi4vc3JjL2Z1bGxCdW5kbGVEZWZhdWx0U3R5bGVPcHRpb25zLnRzIl0sInNvdXJjZXNDb250ZW50IjpbImltcG9ydCB7IGRlZmF1bHRTdHlsZU9wdGlvbnMgfSBmcm9tICdib3RmcmFtZXdvcmstd2ViY2hhdC1hcGknO1xuaW1wb3J0IGFkYXB0aXZlQ2FyZHNEZWZhdWx0U3R5bGVPcHRpb25zIGZyb20gJy4vYWRhcHRpdmVDYXJkcy9kZWZhdWx0U3R5bGVPcHRpb25zJztcblxuaW1wb3J0IEZ1bGxCdW5kbGVTdHlsZU9wdGlvbnMgZnJvbSAnLi90eXBlcy9GdWxsQnVuZGxlU3R5bGVPcHRpb25zJztcblxuY29uc3QgZnVsbEJ1bmRsZURlZmF1bHRTdHlsZU9wdGlvbnM6IFJlcXVpcmVkPEZ1bGxCdW5kbGVTdHlsZU9wdGlvbnM+ID0ge1xuICAuLi5kZWZhdWx0U3R5bGVPcHRpb25zLFxuICAuLi5hZGFwdGl2ZUNhcmRzRGVmYXVsdFN0eWxlT3B0aW9uc1xufTtcblxuZXhwb3J0IGRlZmF1bHQgZnVsbEJ1bmRsZURlZmF1bHRTdHlsZU9wdGlvbnM7XG4iXSwibWFwcGluZ3MiOiI7Ozs7Ozs7Ozs7O0FBQUE7O0FBQ0E7Ozs7OztBQUlBLElBQU1BLDZCQUErRCxtQ0FDaEVDLDJDQURnRSxHQUVoRUMsNEJBRmdFLENBQXJFOztlQUtlRiw2QiJ9