"use strict";

var _interopRequireDefault = require("@babel/runtime/helpers/interopRequireDefault");

var _typeof = require("@babel/runtime/helpers/typeof");

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.default = void 0;

var _slicedToArray2 = _interopRequireDefault(require("@babel/runtime/helpers/slicedToArray"));

var _propTypes = _interopRequireDefault(require("prop-types"));

var _react = _interopRequireWildcard(require("react"));

var _botframeworkWebchatComponent = require("botframework-webchat-component");

var _AdaptiveCardBuilder = _interopRequireDefault(require("./AdaptiveCardBuilder"));

var _AdaptiveCardRenderer = _interopRequireDefault(require("./AdaptiveCardRenderer"));

var _useAdaptiveCardsPackage = _interopRequireDefault(require("../hooks/useAdaptiveCardsPackage"));

var _useStyleOptions3 = _interopRequireDefault(require("../../hooks/useStyleOptions"));

function _getRequireWildcardCache(nodeInterop) { if (typeof WeakMap !== "function") return null; var cacheBabelInterop = new WeakMap(); var cacheNodeInterop = new WeakMap(); return (_getRequireWildcardCache = function _getRequireWildcardCache(nodeInterop) { return nodeInterop ? cacheNodeInterop : cacheBabelInterop; })(nodeInterop); }

function _interopRequireWildcard(obj, nodeInterop) { if (!nodeInterop && obj && obj.__esModule) { return obj; } if (obj === null || _typeof(obj) !== "object" && typeof obj !== "function") { return { default: obj }; } var cache = _getRequireWildcardCache(nodeInterop); if (cache && cache.has(obj)) { return cache.get(obj); } var newObj = {}; var hasPropertyDescriptor = Object.defineProperty && Object.getOwnPropertyDescriptor; for (var key in obj) { if (key !== "default" && Object.prototype.hasOwnProperty.call(obj, key)) { var desc = hasPropertyDescriptor ? Object.getOwnPropertyDescriptor(obj, key) : null; if (desc && (desc.get || desc.set)) { Object.defineProperty(newObj, key, desc); } else { newObj[key] = obj[key]; } } } newObj.default = obj; if (cache) { cache.set(obj, newObj); } return newObj; }

var useDirection = _botframeworkWebchatComponent.hooks.useDirection;

var CommonCard = function CommonCard(_ref) {
  var actionPerformedClassName = _ref.actionPerformedClassName,
      content = _ref.content,
      disabled = _ref.disabled;

  var _useAdaptiveCardsPack = (0, _useAdaptiveCardsPackage.default)(),
      _useAdaptiveCardsPack2 = (0, _slicedToArray2.default)(_useAdaptiveCardsPack, 1),
      adaptiveCardsPackage = _useAdaptiveCardsPack2[0];

  var _useDirection = useDirection(),
      _useDirection2 = (0, _slicedToArray2.default)(_useDirection, 1),
      direction = _useDirection2[0];

  var _useStyleOptions = (0, _useStyleOptions3.default)(),
      _useStyleOptions2 = (0, _slicedToArray2.default)(_useStyleOptions, 1),
      styleOptions = _useStyleOptions2[0];

  var builtCard = (0, _react.useMemo)(function () {
    if (content) {
      var builder = new _AdaptiveCardBuilder.default(adaptiveCardsPackage, styleOptions, direction);
      builder.addCommon(content);
      return builder.card;
    }
  }, [adaptiveCardsPackage, content, direction, styleOptions]);
  return /*#__PURE__*/_react.default.createElement(_AdaptiveCardRenderer.default, {
    actionPerformedClassName: actionPerformedClassName,
    adaptiveCard: builtCard,
    disabled: disabled,
    tapAction: content && content.tap
  });
};

CommonCard.defaultProps = {
  actionPerformedClassName: '',
  disabled: undefined
};
CommonCard.propTypes = {
  actionPerformedClassName: _propTypes.default.string,
  content: _propTypes.default.shape({
    tap: _propTypes.default.any
  }).isRequired,
  disabled: _propTypes.default.bool
};
var _default = CommonCard;
exports.default = _default;
//# sourceMappingURL=data:application/json;charset=utf-8;base64,eyJ2ZXJzaW9uIjozLCJuYW1lcyI6WyJ1c2VEaXJlY3Rpb24iLCJob29rcyIsIkNvbW1vbkNhcmQiLCJhY3Rpb25QZXJmb3JtZWRDbGFzc05hbWUiLCJjb250ZW50IiwiZGlzYWJsZWQiLCJ1c2VBZGFwdGl2ZUNhcmRzUGFja2FnZSIsImFkYXB0aXZlQ2FyZHNQYWNrYWdlIiwiZGlyZWN0aW9uIiwidXNlU3R5bGVPcHRpb25zIiwic3R5bGVPcHRpb25zIiwiYnVpbHRDYXJkIiwidXNlTWVtbyIsImJ1aWxkZXIiLCJBZGFwdGl2ZUNhcmRCdWlsZGVyIiwiYWRkQ29tbW9uIiwiY2FyZCIsInRhcCIsImRlZmF1bHRQcm9wcyIsInVuZGVmaW5lZCIsInByb3BUeXBlcyIsIlByb3BUeXBlcyIsInN0cmluZyIsInNoYXBlIiwiYW55IiwiaXNSZXF1aXJlZCIsImJvb2wiXSwic291cmNlUm9vdCI6ImJ1bmRsZTovLy8iLCJzb3VyY2VzIjpbIi4uLy4uLy4uL3NyYy9hZGFwdGl2ZUNhcmRzL0F0dGFjaG1lbnQvQ29tbW9uQ2FyZC5qcyJdLCJzb3VyY2VzQ29udGVudCI6WyJpbXBvcnQgUHJvcFR5cGVzIGZyb20gJ3Byb3AtdHlwZXMnO1xuaW1wb3J0IFJlYWN0LCB7IHVzZU1lbW8gfSBmcm9tICdyZWFjdCc7XG5cbmltcG9ydCB7IGhvb2tzIH0gZnJvbSAnYm90ZnJhbWV3b3JrLXdlYmNoYXQtY29tcG9uZW50JztcblxuaW1wb3J0IEFkYXB0aXZlQ2FyZEJ1aWxkZXIgZnJvbSAnLi9BZGFwdGl2ZUNhcmRCdWlsZGVyJztcbmltcG9ydCBBZGFwdGl2ZUNhcmRSZW5kZXJlciBmcm9tICcuL0FkYXB0aXZlQ2FyZFJlbmRlcmVyJztcbmltcG9ydCB1c2VBZGFwdGl2ZUNhcmRzUGFja2FnZSBmcm9tICcuLi9ob29rcy91c2VBZGFwdGl2ZUNhcmRzUGFja2FnZSc7XG5pbXBvcnQgdXNlU3R5bGVPcHRpb25zIGZyb20gJy4uLy4uL2hvb2tzL3VzZVN0eWxlT3B0aW9ucyc7XG5cbmNvbnN0IHsgdXNlRGlyZWN0aW9uIH0gPSBob29rcztcblxuY29uc3QgQ29tbW9uQ2FyZCA9ICh7IGFjdGlvblBlcmZvcm1lZENsYXNzTmFtZSwgY29udGVudCwgZGlzYWJsZWQgfSkgPT4ge1xuICBjb25zdCBbYWRhcHRpdmVDYXJkc1BhY2thZ2VdID0gdXNlQWRhcHRpdmVDYXJkc1BhY2thZ2UoKTtcbiAgY29uc3QgW2RpcmVjdGlvbl0gPSB1c2VEaXJlY3Rpb24oKTtcbiAgY29uc3QgW3N0eWxlT3B0aW9uc10gPSB1c2VTdHlsZU9wdGlvbnMoKTtcblxuICBjb25zdCBidWlsdENhcmQgPSB1c2VNZW1vKCgpID0+IHtcbiAgICBpZiAoY29udGVudCkge1xuICAgICAgY29uc3QgYnVpbGRlciA9IG5ldyBBZGFwdGl2ZUNhcmRCdWlsZGVyKGFkYXB0aXZlQ2FyZHNQYWNrYWdlLCBzdHlsZU9wdGlvbnMsIGRpcmVjdGlvbik7XG5cbiAgICAgIGJ1aWxkZXIuYWRkQ29tbW9uKGNvbnRlbnQpO1xuXG4gICAgICByZXR1cm4gYnVpbGRlci5jYXJkO1xuICAgIH1cbiAgfSwgW2FkYXB0aXZlQ2FyZHNQYWNrYWdlLCBjb250ZW50LCBkaXJlY3Rpb24sIHN0eWxlT3B0aW9uc10pO1xuXG4gIHJldHVybiAoXG4gICAgPEFkYXB0aXZlQ2FyZFJlbmRlcmVyXG4gICAgICBhY3Rpb25QZXJmb3JtZWRDbGFzc05hbWU9e2FjdGlvblBlcmZvcm1lZENsYXNzTmFtZX1cbiAgICAgIGFkYXB0aXZlQ2FyZD17YnVpbHRDYXJkfVxuICAgICAgZGlzYWJsZWQ9e2Rpc2FibGVkfVxuICAgICAgdGFwQWN0aW9uPXtjb250ZW50ICYmIGNvbnRlbnQudGFwfVxuICAgIC8+XG4gICk7XG59O1xuXG5Db21tb25DYXJkLmRlZmF1bHRQcm9wcyA9IHtcbiAgYWN0aW9uUGVyZm9ybWVkQ2xhc3NOYW1lOiAnJyxcbiAgZGlzYWJsZWQ6IHVuZGVmaW5lZFxufTtcblxuQ29tbW9uQ2FyZC5wcm9wVHlwZXMgPSB7XG4gIGFjdGlvblBlcmZvcm1lZENsYXNzTmFtZTogUHJvcFR5cGVzLnN0cmluZyxcbiAgY29udGVudDogUHJvcFR5cGVzLnNoYXBlKHtcbiAgICB0YXA6IFByb3BUeXBlcy5hbnlcbiAgfSkuaXNSZXF1aXJlZCxcbiAgZGlzYWJsZWQ6IFByb3BUeXBlcy5ib29sXG59O1xuXG5leHBvcnQgZGVmYXVsdCBDb21tb25DYXJkO1xuIl0sIm1hcHBpbmdzIjoiOzs7Ozs7Ozs7Ozs7O0FBQUE7O0FBQ0E7O0FBRUE7O0FBRUE7O0FBQ0E7O0FBQ0E7O0FBQ0E7Ozs7OztBQUVBLElBQVFBLFlBQVIsR0FBeUJDLG1DQUF6QixDQUFRRCxZQUFSOztBQUVBLElBQU1FLFVBQVUsR0FBRyxTQUFiQSxVQUFhLE9BQXFEO0VBQUEsSUFBbERDLHdCQUFrRCxRQUFsREEsd0JBQWtEO0VBQUEsSUFBeEJDLE9BQXdCLFFBQXhCQSxPQUF3QjtFQUFBLElBQWZDLFFBQWUsUUFBZkEsUUFBZTs7RUFDdEUsNEJBQStCLElBQUFDLGdDQUFBLEdBQS9CO0VBQUE7RUFBQSxJQUFPQyxvQkFBUDs7RUFDQSxvQkFBb0JQLFlBQVksRUFBaEM7RUFBQTtFQUFBLElBQU9RLFNBQVA7O0VBQ0EsdUJBQXVCLElBQUFDLHlCQUFBLEdBQXZCO0VBQUE7RUFBQSxJQUFPQyxZQUFQOztFQUVBLElBQU1DLFNBQVMsR0FBRyxJQUFBQyxjQUFBLEVBQVEsWUFBTTtJQUM5QixJQUFJUixPQUFKLEVBQWE7TUFDWCxJQUFNUyxPQUFPLEdBQUcsSUFBSUMsNEJBQUosQ0FBd0JQLG9CQUF4QixFQUE4Q0csWUFBOUMsRUFBNERGLFNBQTVELENBQWhCO01BRUFLLE9BQU8sQ0FBQ0UsU0FBUixDQUFrQlgsT0FBbEI7TUFFQSxPQUFPUyxPQUFPLENBQUNHLElBQWY7SUFDRDtFQUNGLENBUmlCLEVBUWYsQ0FBQ1Qsb0JBQUQsRUFBdUJILE9BQXZCLEVBQWdDSSxTQUFoQyxFQUEyQ0UsWUFBM0MsQ0FSZSxDQUFsQjtFQVVBLG9CQUNFLDZCQUFDLDZCQUFEO0lBQ0Usd0JBQXdCLEVBQUVQLHdCQUQ1QjtJQUVFLFlBQVksRUFBRVEsU0FGaEI7SUFHRSxRQUFRLEVBQUVOLFFBSFo7SUFJRSxTQUFTLEVBQUVELE9BQU8sSUFBSUEsT0FBTyxDQUFDYTtFQUpoQyxFQURGO0FBUUQsQ0F2QkQ7O0FBeUJBZixVQUFVLENBQUNnQixZQUFYLEdBQTBCO0VBQ3hCZix3QkFBd0IsRUFBRSxFQURGO0VBRXhCRSxRQUFRLEVBQUVjO0FBRmMsQ0FBMUI7QUFLQWpCLFVBQVUsQ0FBQ2tCLFNBQVgsR0FBdUI7RUFDckJqQix3QkFBd0IsRUFBRWtCLGtCQUFBLENBQVVDLE1BRGY7RUFFckJsQixPQUFPLEVBQUVpQixrQkFBQSxDQUFVRSxLQUFWLENBQWdCO0lBQ3ZCTixHQUFHLEVBQUVJLGtCQUFBLENBQVVHO0VBRFEsQ0FBaEIsRUFFTkMsVUFKa0I7RUFLckJwQixRQUFRLEVBQUVnQixrQkFBQSxDQUFVSztBQUxDLENBQXZCO2VBUWV4QixVIn0=