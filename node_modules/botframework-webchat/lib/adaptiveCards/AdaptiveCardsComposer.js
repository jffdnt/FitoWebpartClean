"use strict";

var _interopRequireDefault = require("@babel/runtime/helpers/interopRequireDefault");

var _typeof = require("@babel/runtime/helpers/typeof");

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.default = void 0;

var defaultAdaptiveCardsPackage = _interopRequireWildcard(require("adaptivecards"));

var _propTypes = _interopRequireDefault(require("prop-types"));

var _react = _interopRequireWildcard(require("react"));

var _AdaptiveCardsContext = _interopRequireDefault(require("./AdaptiveCardsContext"));

function _getRequireWildcardCache(nodeInterop) { if (typeof WeakMap !== "function") return null; var cacheBabelInterop = new WeakMap(); var cacheNodeInterop = new WeakMap(); return (_getRequireWildcardCache = function _getRequireWildcardCache(nodeInterop) { return nodeInterop ? cacheNodeInterop : cacheBabelInterop; })(nodeInterop); }

function _interopRequireWildcard(obj, nodeInterop) { if (!nodeInterop && obj && obj.__esModule) { return obj; } if (obj === null || _typeof(obj) !== "object" && typeof obj !== "function") { return { default: obj }; } var cache = _getRequireWildcardCache(nodeInterop); if (cache && cache.has(obj)) { return cache.get(obj); } var newObj = {}; var hasPropertyDescriptor = Object.defineProperty && Object.getOwnPropertyDescriptor; for (var key in obj) { if (key !== "default" && Object.prototype.hasOwnProperty.call(obj, key)) { var desc = hasPropertyDescriptor ? Object.getOwnPropertyDescriptor(obj, key) : null; if (desc && (desc.get || desc.set)) { Object.defineProperty(newObj, key, desc); } else { newObj[key] = obj[key]; } } } newObj.default = obj; if (cache) { cache.set(obj, newObj); } return newObj; }

var AdaptiveCardsComposer = function AdaptiveCardsComposer(_ref) {
  var adaptiveCardsHostConfig = _ref.adaptiveCardsHostConfig,
      adaptiveCardsPackage = _ref.adaptiveCardsPackage,
      children = _ref.children;
  var patchedAdaptiveCardsPackage = (0, _react.useMemo)(function () {
    return adaptiveCardsPackage || defaultAdaptiveCardsPackage;
  }, [adaptiveCardsPackage]);
  var adaptiveCardsContext = (0, _react.useMemo)(function () {
    return {
      adaptiveCardsPackage: patchedAdaptiveCardsPackage,
      hostConfigFromProps: adaptiveCardsHostConfig
    };
  }, [adaptiveCardsHostConfig, patchedAdaptiveCardsPackage]);
  return /*#__PURE__*/_react.default.createElement(_AdaptiveCardsContext.default.Provider, {
    value: adaptiveCardsContext
  }, children);
};

AdaptiveCardsComposer.defaultProps = {
  adaptiveCardsHostConfig: undefined,
  adaptiveCardsPackage: undefined,
  children: undefined
};
AdaptiveCardsComposer.propTypes = {
  adaptiveCardsHostConfig: _propTypes.default.any,
  adaptiveCardsPackage: _propTypes.default.any,
  children: _propTypes.default.any
};
var _default = AdaptiveCardsComposer;
exports.default = _default;
//# sourceMappingURL=data:application/json;charset=utf-8;base64,eyJ2ZXJzaW9uIjozLCJuYW1lcyI6WyJBZGFwdGl2ZUNhcmRzQ29tcG9zZXIiLCJhZGFwdGl2ZUNhcmRzSG9zdENvbmZpZyIsImFkYXB0aXZlQ2FyZHNQYWNrYWdlIiwiY2hpbGRyZW4iLCJwYXRjaGVkQWRhcHRpdmVDYXJkc1BhY2thZ2UiLCJ1c2VNZW1vIiwiZGVmYXVsdEFkYXB0aXZlQ2FyZHNQYWNrYWdlIiwiYWRhcHRpdmVDYXJkc0NvbnRleHQiLCJob3N0Q29uZmlnRnJvbVByb3BzIiwiZGVmYXVsdFByb3BzIiwidW5kZWZpbmVkIiwicHJvcFR5cGVzIiwiUHJvcFR5cGVzIiwiYW55Il0sInNvdXJjZVJvb3QiOiJidW5kbGU6Ly8vIiwic291cmNlcyI6WyIuLi8uLi9zcmMvYWRhcHRpdmVDYXJkcy9BZGFwdGl2ZUNhcmRzQ29tcG9zZXIudHN4Il0sInNvdXJjZXNDb250ZW50IjpbImltcG9ydCAqIGFzIGRlZmF1bHRBZGFwdGl2ZUNhcmRzUGFja2FnZSBmcm9tICdhZGFwdGl2ZWNhcmRzJztcbmltcG9ydCBQcm9wVHlwZXMgZnJvbSAncHJvcC10eXBlcyc7XG5pbXBvcnQgUmVhY3QsIHsgRkMsIFJlYWN0Tm9kZSwgdXNlTWVtbyB9IGZyb20gJ3JlYWN0JztcblxuaW1wb3J0IEFkYXB0aXZlQ2FyZHNDb250ZXh0IGZyb20gJy4vQWRhcHRpdmVDYXJkc0NvbnRleHQnO1xuaW1wb3J0IEFkYXB0aXZlQ2FyZHNQYWNrYWdlIGZyb20gJy4uL3R5cGVzL0FkYXB0aXZlQ2FyZHNQYWNrYWdlJztcblxudHlwZSBBZGFwdGl2ZUNhcmRzQ29tcG9zZXJQcm9wcyA9IHtcbiAgYWRhcHRpdmVDYXJkc0hvc3RDb25maWc6IGFueTtcbiAgYWRhcHRpdmVDYXJkc1BhY2thZ2U6IEFkYXB0aXZlQ2FyZHNQYWNrYWdlO1xuICBjaGlsZHJlbjogUmVhY3ROb2RlO1xufTtcblxuY29uc3QgQWRhcHRpdmVDYXJkc0NvbXBvc2VyOiBGQzxBZGFwdGl2ZUNhcmRzQ29tcG9zZXJQcm9wcz4gPSAoe1xuICBhZGFwdGl2ZUNhcmRzSG9zdENvbmZpZyxcbiAgYWRhcHRpdmVDYXJkc1BhY2thZ2UsXG4gIGNoaWxkcmVuXG59KSA9PiB7XG4gIGNvbnN0IHBhdGNoZWRBZGFwdGl2ZUNhcmRzUGFja2FnZSA9IHVzZU1lbW8oXG4gICAgKCkgPT4gYWRhcHRpdmVDYXJkc1BhY2thZ2UgfHwgZGVmYXVsdEFkYXB0aXZlQ2FyZHNQYWNrYWdlLFxuICAgIFthZGFwdGl2ZUNhcmRzUGFja2FnZV1cbiAgKTtcblxuICBjb25zdCBhZGFwdGl2ZUNhcmRzQ29udGV4dCA9IHVzZU1lbW8oXG4gICAgKCkgPT4gKHtcbiAgICAgIGFkYXB0aXZlQ2FyZHNQYWNrYWdlOiBwYXRjaGVkQWRhcHRpdmVDYXJkc1BhY2thZ2UsXG4gICAgICBob3N0Q29uZmlnRnJvbVByb3BzOiBhZGFwdGl2ZUNhcmRzSG9zdENvbmZpZ1xuICAgIH0pLFxuICAgIFthZGFwdGl2ZUNhcmRzSG9zdENvbmZpZywgcGF0Y2hlZEFkYXB0aXZlQ2FyZHNQYWNrYWdlXVxuICApO1xuXG4gIHJldHVybiA8QWRhcHRpdmVDYXJkc0NvbnRleHQuUHJvdmlkZXIgdmFsdWU9e2FkYXB0aXZlQ2FyZHNDb250ZXh0fT57Y2hpbGRyZW59PC9BZGFwdGl2ZUNhcmRzQ29udGV4dC5Qcm92aWRlcj47XG59O1xuXG5BZGFwdGl2ZUNhcmRzQ29tcG9zZXIuZGVmYXVsdFByb3BzID0ge1xuICBhZGFwdGl2ZUNhcmRzSG9zdENvbmZpZzogdW5kZWZpbmVkLFxuICBhZGFwdGl2ZUNhcmRzUGFja2FnZTogdW5kZWZpbmVkLFxuICBjaGlsZHJlbjogdW5kZWZpbmVkXG59O1xuXG5BZGFwdGl2ZUNhcmRzQ29tcG9zZXIucHJvcFR5cGVzID0ge1xuICBhZGFwdGl2ZUNhcmRzSG9zdENvbmZpZzogUHJvcFR5cGVzLmFueSxcbiAgYWRhcHRpdmVDYXJkc1BhY2thZ2U6IFByb3BUeXBlcy5hbnksXG4gIGNoaWxkcmVuOiBQcm9wVHlwZXMuYW55XG59O1xuXG5leHBvcnQgZGVmYXVsdCBBZGFwdGl2ZUNhcmRzQ29tcG9zZXI7XG4iXSwibWFwcGluZ3MiOiI7Ozs7Ozs7Ozs7O0FBQUE7O0FBQ0E7O0FBQ0E7O0FBRUE7Ozs7OztBQVNBLElBQU1BLHFCQUFxRCxHQUFHLFNBQXhEQSxxQkFBd0QsT0FJeEQ7RUFBQSxJQUhKQyx1QkFHSSxRQUhKQSx1QkFHSTtFQUFBLElBRkpDLG9CQUVJLFFBRkpBLG9CQUVJO0VBQUEsSUFESkMsUUFDSSxRQURKQSxRQUNJO0VBQ0osSUFBTUMsMkJBQTJCLEdBQUcsSUFBQUMsY0FBQSxFQUNsQztJQUFBLE9BQU1ILG9CQUFvQixJQUFJSSwyQkFBOUI7RUFBQSxDQURrQyxFQUVsQyxDQUFDSixvQkFBRCxDQUZrQyxDQUFwQztFQUtBLElBQU1LLG9CQUFvQixHQUFHLElBQUFGLGNBQUEsRUFDM0I7SUFBQSxPQUFPO01BQ0xILG9CQUFvQixFQUFFRSwyQkFEakI7TUFFTEksbUJBQW1CLEVBQUVQO0lBRmhCLENBQVA7RUFBQSxDQUQyQixFQUszQixDQUFDQSx1QkFBRCxFQUEwQkcsMkJBQTFCLENBTDJCLENBQTdCO0VBUUEsb0JBQU8sNkJBQUMsNkJBQUQsQ0FBc0IsUUFBdEI7SUFBK0IsS0FBSyxFQUFFRztFQUF0QyxHQUE2REosUUFBN0QsQ0FBUDtBQUNELENBbkJEOztBQXFCQUgscUJBQXFCLENBQUNTLFlBQXRCLEdBQXFDO0VBQ25DUix1QkFBdUIsRUFBRVMsU0FEVTtFQUVuQ1Isb0JBQW9CLEVBQUVRLFNBRmE7RUFHbkNQLFFBQVEsRUFBRU87QUFIeUIsQ0FBckM7QUFNQVYscUJBQXFCLENBQUNXLFNBQXRCLEdBQWtDO0VBQ2hDVix1QkFBdUIsRUFBRVcsa0JBQUEsQ0FBVUMsR0FESDtFQUVoQ1gsb0JBQW9CLEVBQUVVLGtCQUFBLENBQVVDLEdBRkE7RUFHaENWLFFBQVEsRUFBRVMsa0JBQUEsQ0FBVUM7QUFIWSxDQUFsQztlQU1lYixxQiJ9