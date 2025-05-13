"use strict";

var _interopRequireDefault = require("@babel/runtime/helpers/interopRequireDefault");

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.default = void 0;

var _defineProperty2 = _interopRequireDefault(require("@babel/runtime/helpers/defineProperty"));

var _extends2 = _interopRequireDefault(require("@babel/runtime/helpers/extends"));

var _botframeworkWebchatComponent = require("botframework-webchat-component");

var _propTypes = _interopRequireDefault(require("prop-types"));

var _react = _interopRequireDefault(require("react"));

var _AddFullBundle = _interopRequireDefault(require("./AddFullBundle"));

function ownKeys(object, enumerableOnly) { var keys = Object.keys(object); if (Object.getOwnPropertySymbols) { var symbols = Object.getOwnPropertySymbols(object); enumerableOnly && (symbols = symbols.filter(function (sym) { return Object.getOwnPropertyDescriptor(object, sym).enumerable; })), keys.push.apply(keys, symbols); } return keys; }

function _objectSpread(target) { for (var i = 1; i < arguments.length; i++) { var source = null != arguments[i] ? arguments[i] : {}; i % 2 ? ownKeys(Object(source), !0).forEach(function (key) { (0, _defineProperty2.default)(target, key, source[key]); }) : Object.getOwnPropertyDescriptors ? Object.defineProperties(target, Object.getOwnPropertyDescriptors(source)) : ownKeys(Object(source)).forEach(function (key) { Object.defineProperty(target, key, Object.getOwnPropertyDescriptor(source, key)); }); } return target; }

var Composer = _botframeworkWebchatComponent.Components.Composer;

var FullComposer = function FullComposer(props) {
  return /*#__PURE__*/_react.default.createElement(_AddFullBundle.default, props, function (extraProps) {
    return /*#__PURE__*/_react.default.createElement(Composer, (0, _extends2.default)({}, props, extraProps), props.children);
  });
};

FullComposer.defaultProps = _objectSpread(_objectSpread({}, Composer.defaultProps), {}, {
  adaptiveCardsHostConfig: undefined,
  adaptiveCardsPackage: undefined,
  children: undefined
});
FullComposer.propTypes = _objectSpread(_objectSpread({}, Composer.propTypes), {}, {
  adaptiveCardsHostConfig: _propTypes.default.any,
  adaptiveCardsPackage: _propTypes.default.any,
  children: _propTypes.default.any
});
var _default = FullComposer;
exports.default = _default;
//# sourceMappingURL=data:application/json;charset=utf-8;base64,eyJ2ZXJzaW9uIjozLCJuYW1lcyI6WyJDb21wb3NlciIsIkNvbXBvbmVudHMiLCJGdWxsQ29tcG9zZXIiLCJwcm9wcyIsImV4dHJhUHJvcHMiLCJjaGlsZHJlbiIsImRlZmF1bHRQcm9wcyIsImFkYXB0aXZlQ2FyZHNIb3N0Q29uZmlnIiwidW5kZWZpbmVkIiwiYWRhcHRpdmVDYXJkc1BhY2thZ2UiLCJwcm9wVHlwZXMiLCJQcm9wVHlwZXMiLCJhbnkiXSwic291cmNlUm9vdCI6ImJ1bmRsZTovLy8iLCJzb3VyY2VzIjpbIi4uL3NyYy9GdWxsQ29tcG9zZXIudHN4Il0sInNvdXJjZXNDb250ZW50IjpbImltcG9ydCB7IENvbXBvbmVudHMgfSBmcm9tICdib3RmcmFtZXdvcmstd2ViY2hhdC1jb21wb25lbnQnO1xuaW1wb3J0IFByb3BUeXBlcyBmcm9tICdwcm9wLXR5cGVzJztcbmltcG9ydCBSZWFjdCBmcm9tICdyZWFjdCc7XG5cbmltcG9ydCBBZGRGdWxsQnVuZGxlIGZyb20gJy4vQWRkRnVsbEJ1bmRsZSc7XG5cbmltcG9ydCB0eXBlIHsgQWRkRnVsbEJ1bmRsZVByb3BzIH0gZnJvbSAnLi9BZGRGdWxsQnVuZGxlJztcbmltcG9ydCB0eXBlIHsgQ29tcG9zZXJQcm9wcyB9IGZyb20gJ2JvdGZyYW1ld29yay13ZWJjaGF0LWNvbXBvbmVudCc7XG5pbXBvcnQgdHlwZSB7IEZDIH0gZnJvbSAncmVhY3QnO1xuXG5jb25zdCB7IENvbXBvc2VyIH0gPSBDb21wb25lbnRzO1xuXG50eXBlIEZ1bGxDb21wb3NlclByb3BzID0gQ29tcG9zZXJQcm9wcyAmIEFkZEZ1bGxCdW5kbGVQcm9wcztcblxuY29uc3QgRnVsbENvbXBvc2VyOiBGQzxGdWxsQ29tcG9zZXJQcm9wcz4gPSBwcm9wcyA9PiAoXG4gIDxBZGRGdWxsQnVuZGxlIHsuLi5wcm9wc30+XG4gICAge2V4dHJhUHJvcHMgPT4gKFxuICAgICAgPENvbXBvc2VyIHsuLi5wcm9wc30gey4uLmV4dHJhUHJvcHN9PlxuICAgICAgICB7LyogV2UgbmVlZCB0byBzcHJlYWQsIHRodXMsIHdlIGNhbm5vdCB3ZSBkZXN0cnVjdHVyaW5nIGFzc2lnbm1lbnQuICovfVxuICAgICAgICB7LyogZXNsaW50LWRpc2FibGUtbmV4dC1saW5lIHJlYWN0L2Rlc3RydWN0dXJpbmctYXNzaWdubWVudCAqL31cbiAgICAgICAge3Byb3BzLmNoaWxkcmVufVxuICAgICAgPC9Db21wb3Nlcj5cbiAgICApfVxuICA8L0FkZEZ1bGxCdW5kbGU+XG4pO1xuXG5GdWxsQ29tcG9zZXIuZGVmYXVsdFByb3BzID0ge1xuICAuLi5Db21wb3Nlci5kZWZhdWx0UHJvcHMsXG4gIGFkYXB0aXZlQ2FyZHNIb3N0Q29uZmlnOiB1bmRlZmluZWQsXG4gIGFkYXB0aXZlQ2FyZHNQYWNrYWdlOiB1bmRlZmluZWQsXG4gIGNoaWxkcmVuOiB1bmRlZmluZWRcbn07XG5cbkZ1bGxDb21wb3Nlci5wcm9wVHlwZXMgPSB7XG4gIC4uLkNvbXBvc2VyLnByb3BUeXBlcyxcbiAgYWRhcHRpdmVDYXJkc0hvc3RDb25maWc6IFByb3BUeXBlcy5hbnksXG4gIGFkYXB0aXZlQ2FyZHNQYWNrYWdlOiBQcm9wVHlwZXMuYW55LFxuICBjaGlsZHJlbjogUHJvcFR5cGVzLmFueVxufTtcblxuZXhwb3J0IGRlZmF1bHQgRnVsbENvbXBvc2VyO1xuXG5leHBvcnQgdHlwZSB7IEZ1bGxDb21wb3NlclByb3BzIH07XG4iXSwibWFwcGluZ3MiOiI7Ozs7Ozs7Ozs7Ozs7QUFBQTs7QUFDQTs7QUFDQTs7QUFFQTs7Ozs7O0FBTUEsSUFBUUEsUUFBUixHQUFxQkMsd0NBQXJCLENBQVFELFFBQVI7O0FBSUEsSUFBTUUsWUFBbUMsR0FBRyxTQUF0Q0EsWUFBc0MsQ0FBQUMsS0FBSztFQUFBLG9CQUMvQyw2QkFBQyxzQkFBRCxFQUFtQkEsS0FBbkIsRUFDRyxVQUFBQyxVQUFVO0lBQUEsb0JBQ1QsNkJBQUMsUUFBRCw2QkFBY0QsS0FBZCxFQUF5QkMsVUFBekIsR0FHR0QsS0FBSyxDQUFDRSxRQUhULENBRFM7RUFBQSxDQURiLENBRCtDO0FBQUEsQ0FBakQ7O0FBWUFILFlBQVksQ0FBQ0ksWUFBYixtQ0FDS04sUUFBUSxDQUFDTSxZQURkO0VBRUVDLHVCQUF1QixFQUFFQyxTQUYzQjtFQUdFQyxvQkFBb0IsRUFBRUQsU0FIeEI7RUFJRUgsUUFBUSxFQUFFRztBQUpaO0FBT0FOLFlBQVksQ0FBQ1EsU0FBYixtQ0FDS1YsUUFBUSxDQUFDVSxTQURkO0VBRUVILHVCQUF1QixFQUFFSSxrQkFBQSxDQUFVQyxHQUZyQztFQUdFSCxvQkFBb0IsRUFBRUUsa0JBQUEsQ0FBVUMsR0FIbEM7RUFJRVAsUUFBUSxFQUFFTSxrQkFBQSxDQUFVQztBQUp0QjtlQU9lVixZIn0=