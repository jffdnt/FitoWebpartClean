"use strict";

var _interopRequireDefault = require("@babel/runtime/helpers/interopRequireDefault");

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.default = void 0;

var _defineProperty2 = _interopRequireDefault(require("@babel/runtime/helpers/defineProperty"));

var _extends2 = _interopRequireDefault(require("@babel/runtime/helpers/extends"));

var _propTypes = _interopRequireDefault(require("prop-types"));

var _react = _interopRequireDefault(require("react"));

var _botframeworkWebchatComponent = _interopRequireDefault(require("botframework-webchat-component"));

var _AddFullBundle = _interopRequireDefault(require("./AddFullBundle"));

function ownKeys(object, enumerableOnly) { var keys = Object.keys(object); if (Object.getOwnPropertySymbols) { var symbols = Object.getOwnPropertySymbols(object); enumerableOnly && (symbols = symbols.filter(function (sym) { return Object.getOwnPropertyDescriptor(object, sym).enumerable; })), keys.push.apply(keys, symbols); } return keys; }

function _objectSpread(target) { for (var i = 1; i < arguments.length; i++) { var source = null != arguments[i] ? arguments[i] : {}; i % 2 ? ownKeys(Object(source), !0).forEach(function (key) { (0, _defineProperty2.default)(target, key, source[key]); }) : Object.getOwnPropertyDescriptors ? Object.defineProperties(target, Object.getOwnPropertyDescriptors(source)) : ownKeys(Object(source)).forEach(function (key) { Object.defineProperty(target, key, Object.getOwnPropertyDescriptor(source, key)); }); } return target; }

// Add additional props to <WebChat>, so it support additional features
var FullReactWebChat = function FullReactWebChat(props) {
  return /*#__PURE__*/_react.default.createElement(_AddFullBundle.default, props, function (extraProps) {
    return /*#__PURE__*/_react.default.createElement(_botframeworkWebchatComponent.default, (0, _extends2.default)({}, props, extraProps));
  });
};

FullReactWebChat.defaultProps = _objectSpread(_objectSpread({}, _botframeworkWebchatComponent.default.defaultProps), {}, {
  adaptiveCardHostConfig: undefined,
  adaptiveCardsHostConfig: undefined,
  adaptiveCardsPackage: undefined,
  renderMarkdown: undefined
});
FullReactWebChat.propTypes = _objectSpread(_objectSpread({}, _botframeworkWebchatComponent.default.propTypes), {}, {
  adaptiveCardHostConfig: _propTypes.default.any,
  adaptiveCardsHostConfig: _propTypes.default.any,
  adaptiveCardsPackage: _propTypes.default.any,
  renderMarkdown: _propTypes.default.func
});
var _default = FullReactWebChat;
exports.default = _default;
//# sourceMappingURL=data:application/json;charset=utf-8;base64,eyJ2ZXJzaW9uIjozLCJuYW1lcyI6WyJGdWxsUmVhY3RXZWJDaGF0IiwicHJvcHMiLCJleHRyYVByb3BzIiwiZGVmYXVsdFByb3BzIiwiUmVhY3RXZWJDaGF0IiwiYWRhcHRpdmVDYXJkSG9zdENvbmZpZyIsInVuZGVmaW5lZCIsImFkYXB0aXZlQ2FyZHNIb3N0Q29uZmlnIiwiYWRhcHRpdmVDYXJkc1BhY2thZ2UiLCJyZW5kZXJNYXJrZG93biIsInByb3BUeXBlcyIsIlByb3BUeXBlcyIsImFueSIsImZ1bmMiXSwic291cmNlUm9vdCI6ImJ1bmRsZTovLy8iLCJzb3VyY2VzIjpbIi4uL3NyYy9GdWxsUmVhY3RXZWJDaGF0LnRzeCJdLCJzb3VyY2VzQ29udGVudCI6WyJpbXBvcnQgUHJvcFR5cGVzIGZyb20gJ3Byb3AtdHlwZXMnO1xuaW1wb3J0IFJlYWN0LCB7IFZGQyB9IGZyb20gJ3JlYWN0JztcbmltcG9ydCBSZWFjdFdlYkNoYXQsIHsgUmVhY3RXZWJDaGF0UHJvcHMgfSBmcm9tICdib3RmcmFtZXdvcmstd2ViY2hhdC1jb21wb25lbnQnO1xuXG5pbXBvcnQgQWRkRnVsbEJ1bmRsZSwgeyBBZGRGdWxsQnVuZGxlUHJvcHMgfSBmcm9tICcuL0FkZEZ1bGxCdW5kbGUnO1xuXG50eXBlIEZ1bGxSZWFjdFdlYkNoYXRQcm9wcyA9IFJlYWN0V2ViQ2hhdFByb3BzICYgT21pdDxBZGRGdWxsQnVuZGxlUHJvcHMsICdjaGlsZHJlbic+O1xuXG4vLyBBZGQgYWRkaXRpb25hbCBwcm9wcyB0byA8V2ViQ2hhdD4sIHNvIGl0IHN1cHBvcnQgYWRkaXRpb25hbCBmZWF0dXJlc1xuY29uc3QgRnVsbFJlYWN0V2ViQ2hhdDogVkZDPEZ1bGxSZWFjdFdlYkNoYXRQcm9wcz4gPSBwcm9wcyA9PiAoXG4gIDxBZGRGdWxsQnVuZGxlIHsuLi5wcm9wc30+e2V4dHJhUHJvcHMgPT4gPFJlYWN0V2ViQ2hhdCB7Li4ucHJvcHN9IHsuLi5leHRyYVByb3BzfSAvPn08L0FkZEZ1bGxCdW5kbGU+XG4pO1xuXG5GdWxsUmVhY3RXZWJDaGF0LmRlZmF1bHRQcm9wcyA9IHtcbiAgLi4uUmVhY3RXZWJDaGF0LmRlZmF1bHRQcm9wcyxcbiAgYWRhcHRpdmVDYXJkSG9zdENvbmZpZzogdW5kZWZpbmVkLFxuICBhZGFwdGl2ZUNhcmRzSG9zdENvbmZpZzogdW5kZWZpbmVkLFxuICBhZGFwdGl2ZUNhcmRzUGFja2FnZTogdW5kZWZpbmVkLFxuICByZW5kZXJNYXJrZG93bjogdW5kZWZpbmVkXG59O1xuXG5GdWxsUmVhY3RXZWJDaGF0LnByb3BUeXBlcyA9IHtcbiAgLi4uUmVhY3RXZWJDaGF0LnByb3BUeXBlcyxcbiAgYWRhcHRpdmVDYXJkSG9zdENvbmZpZzogUHJvcFR5cGVzLmFueSxcbiAgYWRhcHRpdmVDYXJkc0hvc3RDb25maWc6IFByb3BUeXBlcy5hbnksXG4gIGFkYXB0aXZlQ2FyZHNQYWNrYWdlOiBQcm9wVHlwZXMuYW55LFxuICByZW5kZXJNYXJrZG93bjogUHJvcFR5cGVzLmZ1bmNcbn07XG5cbmV4cG9ydCBkZWZhdWx0IEZ1bGxSZWFjdFdlYkNoYXQ7XG5cbmV4cG9ydCB0eXBlIHsgRnVsbFJlYWN0V2ViQ2hhdFByb3BzIH07XG4iXSwibWFwcGluZ3MiOiI7Ozs7Ozs7Ozs7Ozs7QUFBQTs7QUFDQTs7QUFDQTs7QUFFQTs7Ozs7O0FBSUE7QUFDQSxJQUFNQSxnQkFBNEMsR0FBRyxTQUEvQ0EsZ0JBQStDLENBQUFDLEtBQUs7RUFBQSxvQkFDeEQsNkJBQUMsc0JBQUQsRUFBbUJBLEtBQW5CLEVBQTJCLFVBQUFDLFVBQVU7SUFBQSxvQkFBSSw2QkFBQyxxQ0FBRCw2QkFBa0JELEtBQWxCLEVBQTZCQyxVQUE3QixFQUFKO0VBQUEsQ0FBckMsQ0FEd0Q7QUFBQSxDQUExRDs7QUFJQUYsZ0JBQWdCLENBQUNHLFlBQWpCLG1DQUNLQyxxQ0FBQSxDQUFhRCxZQURsQjtFQUVFRSxzQkFBc0IsRUFBRUMsU0FGMUI7RUFHRUMsdUJBQXVCLEVBQUVELFNBSDNCO0VBSUVFLG9CQUFvQixFQUFFRixTQUp4QjtFQUtFRyxjQUFjLEVBQUVIO0FBTGxCO0FBUUFOLGdCQUFnQixDQUFDVSxTQUFqQixtQ0FDS04scUNBQUEsQ0FBYU0sU0FEbEI7RUFFRUwsc0JBQXNCLEVBQUVNLGtCQUFBLENBQVVDLEdBRnBDO0VBR0VMLHVCQUF1QixFQUFFSSxrQkFBQSxDQUFVQyxHQUhyQztFQUlFSixvQkFBb0IsRUFBRUcsa0JBQUEsQ0FBVUMsR0FKbEM7RUFLRUgsY0FBYyxFQUFFRSxrQkFBQSxDQUFVRTtBQUw1QjtlQVFlYixnQiJ9