"use strict";

var _interopRequireDefault = require("@babel/runtime/helpers/interopRequireDefault");

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.default = normalizeStyleOptions;

var _defineProperty2 = _interopRequireDefault(require("@babel/runtime/helpers/defineProperty"));

var _defaultStyleOptions = _interopRequireDefault(require("./defaultStyleOptions"));

function ownKeys(object, enumerableOnly) { var keys = Object.keys(object); if (Object.getOwnPropertySymbols) { var symbols = Object.getOwnPropertySymbols(object); enumerableOnly && (symbols = symbols.filter(function (sym) { return Object.getOwnPropertyDescriptor(object, sym).enumerable; })), keys.push.apply(keys, symbols); } return keys; }

function _objectSpread(target) { for (var i = 1; i < arguments.length; i++) { var source = null != arguments[i] ? arguments[i] : {}; i % 2 ? ownKeys(Object(source), !0).forEach(function (key) { (0, _defineProperty2.default)(target, key, source[key]); }) : Object.getOwnPropertyDescriptors ? Object.defineProperties(target, Object.getOwnPropertyDescriptors(source)) : ownKeys(Object(source)).forEach(function (key) { Object.defineProperty(target, key, Object.getOwnPropertyDescriptor(source, key)); }); } return target; }

function normalizeStyleOptions(styleOptions) {
  return _objectSpread(_objectSpread({}, _defaultStyleOptions.default), styleOptions);
}
//# sourceMappingURL=data:application/json;charset=utf-8;base64,eyJ2ZXJzaW9uIjozLCJuYW1lcyI6WyJub3JtYWxpemVTdHlsZU9wdGlvbnMiLCJzdHlsZU9wdGlvbnMiLCJkZWZhdWx0U3R5bGVPcHRpb25zIl0sInNvdXJjZVJvb3QiOiJidW5kbGU6Ly8vIiwic291cmNlcyI6WyIuLi8uLi9zcmMvYWRhcHRpdmVDYXJkcy9ub3JtYWxpemVTdHlsZU9wdGlvbnMudHMiXSwic291cmNlc0NvbnRlbnQiOlsiaW1wb3J0IEFkYXB0aXZlQ2FyZHNTdHlsZU9wdGlvbnMsIHsgU3RyaWN0QWRhcHRpdmVDYXJkc1N0eWxlT3B0aW9ucyB9IGZyb20gJy4vQWRhcHRpdmVDYXJkc1N0eWxlT3B0aW9ucyc7XG5pbXBvcnQgZGVmYXVsdFN0eWxlT3B0aW9ucyBmcm9tICcuL2RlZmF1bHRTdHlsZU9wdGlvbnMnO1xuXG5leHBvcnQgZGVmYXVsdCBmdW5jdGlvbiBub3JtYWxpemVTdHlsZU9wdGlvbnMoXG4gIHN0eWxlT3B0aW9uczogQWRhcHRpdmVDYXJkc1N0eWxlT3B0aW9uc1xuKTogU3RyaWN0QWRhcHRpdmVDYXJkc1N0eWxlT3B0aW9ucyB7XG4gIHJldHVybiB7IC4uLmRlZmF1bHRTdHlsZU9wdGlvbnMsIC4uLnN0eWxlT3B0aW9ucyB9O1xufVxuIl0sIm1hcHBpbmdzIjoiOzs7Ozs7Ozs7OztBQUNBOzs7Ozs7QUFFZSxTQUFTQSxxQkFBVCxDQUNiQyxZQURhLEVBRW9CO0VBQ2pDLHVDQUFZQyw0QkFBWixHQUFvQ0QsWUFBcEM7QUFDRCJ9