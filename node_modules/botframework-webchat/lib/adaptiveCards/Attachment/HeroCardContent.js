"use strict";

var _interopRequireDefault = require("@babel/runtime/helpers/interopRequireDefault");

var _typeof = require("@babel/runtime/helpers/typeof");

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.default = void 0;

var _slicedToArray2 = _interopRequireDefault(require("@babel/runtime/helpers/slicedToArray"));

var _botframeworkWebchatComponent = require("botframework-webchat-component");

var _propTypes = _interopRequireDefault(require("prop-types"));

var _react = _interopRequireWildcard(require("react"));

var _AdaptiveCardBuilder = _interopRequireDefault(require("./AdaptiveCardBuilder"));

var _AdaptiveCardRenderer = _interopRequireDefault(require("./AdaptiveCardRenderer"));

var _useAdaptiveCardsPackage = _interopRequireDefault(require("../hooks/useAdaptiveCardsPackage"));

var _useStyleOptions3 = _interopRequireDefault(require("../../hooks/useStyleOptions"));

function _getRequireWildcardCache(nodeInterop) { if (typeof WeakMap !== "function") return null; var cacheBabelInterop = new WeakMap(); var cacheNodeInterop = new WeakMap(); return (_getRequireWildcardCache = function _getRequireWildcardCache(nodeInterop) { return nodeInterop ? cacheNodeInterop : cacheBabelInterop; })(nodeInterop); }

function _interopRequireWildcard(obj, nodeInterop) { if (!nodeInterop && obj && obj.__esModule) { return obj; } if (obj === null || _typeof(obj) !== "object" && typeof obj !== "function") { return { default: obj }; } var cache = _getRequireWildcardCache(nodeInterop); if (cache && cache.has(obj)) { return cache.get(obj); } var newObj = {}; var hasPropertyDescriptor = Object.defineProperty && Object.getOwnPropertyDescriptor; for (var key in obj) { if (key !== "default" && Object.prototype.hasOwnProperty.call(obj, key)) { var desc = hasPropertyDescriptor ? Object.getOwnPropertyDescriptor(obj, key) : null; if (desc && (desc.get || desc.set)) { Object.defineProperty(newObj, key, desc); } else { newObj[key] = obj[key]; } } } newObj.default = obj; if (cache) { cache.set(obj, newObj); } return newObj; }

var useDirection = _botframeworkWebchatComponent.hooks.useDirection;

var HeroCardContent = function HeroCardContent(_ref) {
  var actionPerformedClassName = _ref.actionPerformedClassName,
      content = _ref.content,
      disabled = _ref.disabled;

  var _useAdaptiveCardsPack = (0, _useAdaptiveCardsPackage.default)(),
      _useAdaptiveCardsPack2 = (0, _slicedToArray2.default)(_useAdaptiveCardsPack, 1),
      adaptiveCardsPackage = _useAdaptiveCardsPack2[0];

  var _useStyleOptions = (0, _useStyleOptions3.default)(),
      _useStyleOptions2 = (0, _slicedToArray2.default)(_useStyleOptions, 1),
      styleOptions = _useStyleOptions2[0];

  var _useDirection = useDirection(),
      _useDirection2 = (0, _slicedToArray2.default)(_useDirection, 1),
      direction = _useDirection2[0];

  var builtCard = (0, _react.useMemo)(function () {
    var builder = new _AdaptiveCardBuilder.default(adaptiveCardsPackage, styleOptions, direction);

    if (content) {
      (content.images || []).forEach(function (image) {
        return builder.addImage(image.url, null, image.tap, image.alt);
      });
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

HeroCardContent.defaultProps = {
  actionPerformedClassName: '',
  disabled: undefined
};
HeroCardContent.propTypes = {
  actionPerformedClassName: _propTypes.default.string,
  // PropTypes cannot fully capture TypeScript types.
  // @ts-ignore
  content: _propTypes.default.shape({
    images: _propTypes.default.arrayOf(_propTypes.default.shape({
      alt: _propTypes.default.string.isRequired,
      tap: _propTypes.default.any,
      url: _propTypes.default.string.isRequired
    })),
    tap: _propTypes.default.any
  }).isRequired,
  disabled: _propTypes.default.bool
};
var _default = HeroCardContent;
exports.default = _default;
//# sourceMappingURL=data:application/json;charset=utf-8;base64,eyJ2ZXJzaW9uIjozLCJuYW1lcyI6WyJ1c2VEaXJlY3Rpb24iLCJob29rcyIsIkhlcm9DYXJkQ29udGVudCIsImFjdGlvblBlcmZvcm1lZENsYXNzTmFtZSIsImNvbnRlbnQiLCJkaXNhYmxlZCIsInVzZUFkYXB0aXZlQ2FyZHNQYWNrYWdlIiwiYWRhcHRpdmVDYXJkc1BhY2thZ2UiLCJ1c2VTdHlsZU9wdGlvbnMiLCJzdHlsZU9wdGlvbnMiLCJkaXJlY3Rpb24iLCJidWlsdENhcmQiLCJ1c2VNZW1vIiwiYnVpbGRlciIsIkFkYXB0aXZlQ2FyZEJ1aWxkZXIiLCJpbWFnZXMiLCJmb3JFYWNoIiwiaW1hZ2UiLCJhZGRJbWFnZSIsInVybCIsInRhcCIsImFsdCIsImFkZENvbW1vbiIsImNhcmQiLCJkZWZhdWx0UHJvcHMiLCJ1bmRlZmluZWQiLCJwcm9wVHlwZXMiLCJQcm9wVHlwZXMiLCJzdHJpbmciLCJzaGFwZSIsImFycmF5T2YiLCJpc1JlcXVpcmVkIiwiYW55IiwiYm9vbCJdLCJzb3VyY2VSb290IjoiYnVuZGxlOi8vLyIsInNvdXJjZXMiOlsiLi4vLi4vLi4vc3JjL2FkYXB0aXZlQ2FyZHMvQXR0YWNobWVudC9IZXJvQ2FyZENvbnRlbnQudHN4Il0sInNvdXJjZXNDb250ZW50IjpbImltcG9ydCB7IGhvb2tzIH0gZnJvbSAnYm90ZnJhbWV3b3JrLXdlYmNoYXQtY29tcG9uZW50JztcbmltcG9ydCBQcm9wVHlwZXMgZnJvbSAncHJvcC10eXBlcyc7XG5pbXBvcnQgUmVhY3QsIHsgRkMsIHVzZU1lbW8gfSBmcm9tICdyZWFjdCc7XG5pbXBvcnQgdHlwZSB7IERpcmVjdExpbmVIZXJvQ2FyZCB9IGZyb20gJ2JvdGZyYW1ld29yay13ZWJjaGF0LWNvcmUnO1xuXG5pbXBvcnQgQWRhcHRpdmVDYXJkQnVpbGRlciBmcm9tICcuL0FkYXB0aXZlQ2FyZEJ1aWxkZXInO1xuaW1wb3J0IEFkYXB0aXZlQ2FyZFJlbmRlcmVyIGZyb20gJy4vQWRhcHRpdmVDYXJkUmVuZGVyZXInO1xuaW1wb3J0IHVzZUFkYXB0aXZlQ2FyZHNQYWNrYWdlIGZyb20gJy4uL2hvb2tzL3VzZUFkYXB0aXZlQ2FyZHNQYWNrYWdlJztcbmltcG9ydCB1c2VTdHlsZU9wdGlvbnMgZnJvbSAnLi4vLi4vaG9va3MvdXNlU3R5bGVPcHRpb25zJztcblxuY29uc3QgeyB1c2VEaXJlY3Rpb24gfSA9IGhvb2tzO1xuXG50eXBlIEhlcm9DYXJkQ29udGVudFByb3BzID0ge1xuICBhY3Rpb25QZXJmb3JtZWRDbGFzc05hbWU/OiBzdHJpbmc7XG4gIGNvbnRlbnQ6IERpcmVjdExpbmVIZXJvQ2FyZDtcbiAgZGlzYWJsZWQ/OiBib29sZWFuO1xufTtcblxuY29uc3QgSGVyb0NhcmRDb250ZW50OiBGQzxIZXJvQ2FyZENvbnRlbnRQcm9wcz4gPSAoeyBhY3Rpb25QZXJmb3JtZWRDbGFzc05hbWUsIGNvbnRlbnQsIGRpc2FibGVkIH0pID0+IHtcbiAgY29uc3QgW2FkYXB0aXZlQ2FyZHNQYWNrYWdlXSA9IHVzZUFkYXB0aXZlQ2FyZHNQYWNrYWdlKCk7XG4gIGNvbnN0IFtzdHlsZU9wdGlvbnNdID0gdXNlU3R5bGVPcHRpb25zKCk7XG4gIGNvbnN0IFtkaXJlY3Rpb25dID0gdXNlRGlyZWN0aW9uKCk7XG5cbiAgY29uc3QgYnVpbHRDYXJkID0gdXNlTWVtbygoKSA9PiB7XG4gICAgY29uc3QgYnVpbGRlciA9IG5ldyBBZGFwdGl2ZUNhcmRCdWlsZGVyKGFkYXB0aXZlQ2FyZHNQYWNrYWdlLCBzdHlsZU9wdGlvbnMsIGRpcmVjdGlvbik7XG5cbiAgICBpZiAoY29udGVudCkge1xuICAgICAgKGNvbnRlbnQuaW1hZ2VzIHx8IFtdKS5mb3JFYWNoKGltYWdlID0+IGJ1aWxkZXIuYWRkSW1hZ2UoaW1hZ2UudXJsLCBudWxsLCBpbWFnZS50YXAsIGltYWdlLmFsdCkpO1xuXG4gICAgICBidWlsZGVyLmFkZENvbW1vbihjb250ZW50KTtcblxuICAgICAgcmV0dXJuIGJ1aWxkZXIuY2FyZDtcbiAgICB9XG4gIH0sIFthZGFwdGl2ZUNhcmRzUGFja2FnZSwgY29udGVudCwgZGlyZWN0aW9uLCBzdHlsZU9wdGlvbnNdKTtcblxuICByZXR1cm4gKFxuICAgIDxBZGFwdGl2ZUNhcmRSZW5kZXJlclxuICAgICAgYWN0aW9uUGVyZm9ybWVkQ2xhc3NOYW1lPXthY3Rpb25QZXJmb3JtZWRDbGFzc05hbWV9XG4gICAgICBhZGFwdGl2ZUNhcmQ9e2J1aWx0Q2FyZH1cbiAgICAgIGRpc2FibGVkPXtkaXNhYmxlZH1cbiAgICAgIHRhcEFjdGlvbj17Y29udGVudCAmJiBjb250ZW50LnRhcH1cbiAgICAvPlxuICApO1xufTtcblxuSGVyb0NhcmRDb250ZW50LmRlZmF1bHRQcm9wcyA9IHtcbiAgYWN0aW9uUGVyZm9ybWVkQ2xhc3NOYW1lOiAnJyxcbiAgZGlzYWJsZWQ6IHVuZGVmaW5lZFxufTtcblxuSGVyb0NhcmRDb250ZW50LnByb3BUeXBlcyA9IHtcbiAgYWN0aW9uUGVyZm9ybWVkQ2xhc3NOYW1lOiBQcm9wVHlwZXMuc3RyaW5nLFxuICAvLyBQcm9wVHlwZXMgY2Fubm90IGZ1bGx5IGNhcHR1cmUgVHlwZVNjcmlwdCB0eXBlcy5cbiAgLy8gQHRzLWlnbm9yZVxuICBjb250ZW50OiBQcm9wVHlwZXMuc2hhcGUoe1xuICAgIGltYWdlczogUHJvcFR5cGVzLmFycmF5T2YoXG4gICAgICBQcm9wVHlwZXMuc2hhcGUoe1xuICAgICAgICBhbHQ6IFByb3BUeXBlcy5zdHJpbmcuaXNSZXF1aXJlZCxcbiAgICAgICAgdGFwOiBQcm9wVHlwZXMuYW55LFxuICAgICAgICB1cmw6IFByb3BUeXBlcy5zdHJpbmcuaXNSZXF1aXJlZFxuICAgICAgfSlcbiAgICApLFxuICAgIHRhcDogUHJvcFR5cGVzLmFueVxuICB9KS5pc1JlcXVpcmVkLFxuICBkaXNhYmxlZDogUHJvcFR5cGVzLmJvb2xcbn07XG5cbmV4cG9ydCBkZWZhdWx0IEhlcm9DYXJkQ29udGVudDtcbiJdLCJtYXBwaW5ncyI6Ijs7Ozs7Ozs7Ozs7OztBQUFBOztBQUNBOztBQUNBOztBQUdBOztBQUNBOztBQUNBOztBQUNBOzs7Ozs7QUFFQSxJQUFRQSxZQUFSLEdBQXlCQyxtQ0FBekIsQ0FBUUQsWUFBUjs7QUFRQSxJQUFNRSxlQUF5QyxHQUFHLFNBQTVDQSxlQUE0QyxPQUFxRDtFQUFBLElBQWxEQyx3QkFBa0QsUUFBbERBLHdCQUFrRDtFQUFBLElBQXhCQyxPQUF3QixRQUF4QkEsT0FBd0I7RUFBQSxJQUFmQyxRQUFlLFFBQWZBLFFBQWU7O0VBQ3JHLDRCQUErQixJQUFBQyxnQ0FBQSxHQUEvQjtFQUFBO0VBQUEsSUFBT0Msb0JBQVA7O0VBQ0EsdUJBQXVCLElBQUFDLHlCQUFBLEdBQXZCO0VBQUE7RUFBQSxJQUFPQyxZQUFQOztFQUNBLG9CQUFvQlQsWUFBWSxFQUFoQztFQUFBO0VBQUEsSUFBT1UsU0FBUDs7RUFFQSxJQUFNQyxTQUFTLEdBQUcsSUFBQUMsY0FBQSxFQUFRLFlBQU07SUFDOUIsSUFBTUMsT0FBTyxHQUFHLElBQUlDLDRCQUFKLENBQXdCUCxvQkFBeEIsRUFBOENFLFlBQTlDLEVBQTREQyxTQUE1RCxDQUFoQjs7SUFFQSxJQUFJTixPQUFKLEVBQWE7TUFDWCxDQUFDQSxPQUFPLENBQUNXLE1BQVIsSUFBa0IsRUFBbkIsRUFBdUJDLE9BQXZCLENBQStCLFVBQUFDLEtBQUs7UUFBQSxPQUFJSixPQUFPLENBQUNLLFFBQVIsQ0FBaUJELEtBQUssQ0FBQ0UsR0FBdkIsRUFBNEIsSUFBNUIsRUFBa0NGLEtBQUssQ0FBQ0csR0FBeEMsRUFBNkNILEtBQUssQ0FBQ0ksR0FBbkQsQ0FBSjtNQUFBLENBQXBDO01BRUFSLE9BQU8sQ0FBQ1MsU0FBUixDQUFrQmxCLE9BQWxCO01BRUEsT0FBT1MsT0FBTyxDQUFDVSxJQUFmO0lBQ0Q7RUFDRixDQVZpQixFQVVmLENBQUNoQixvQkFBRCxFQUF1QkgsT0FBdkIsRUFBZ0NNLFNBQWhDLEVBQTJDRCxZQUEzQyxDQVZlLENBQWxCO0VBWUEsb0JBQ0UsNkJBQUMsNkJBQUQ7SUFDRSx3QkFBd0IsRUFBRU4sd0JBRDVCO0lBRUUsWUFBWSxFQUFFUSxTQUZoQjtJQUdFLFFBQVEsRUFBRU4sUUFIWjtJQUlFLFNBQVMsRUFBRUQsT0FBTyxJQUFJQSxPQUFPLENBQUNnQjtFQUpoQyxFQURGO0FBUUQsQ0F6QkQ7O0FBMkJBbEIsZUFBZSxDQUFDc0IsWUFBaEIsR0FBK0I7RUFDN0JyQix3QkFBd0IsRUFBRSxFQURHO0VBRTdCRSxRQUFRLEVBQUVvQjtBQUZtQixDQUEvQjtBQUtBdkIsZUFBZSxDQUFDd0IsU0FBaEIsR0FBNEI7RUFDMUJ2Qix3QkFBd0IsRUFBRXdCLGtCQUFBLENBQVVDLE1BRFY7RUFFMUI7RUFDQTtFQUNBeEIsT0FBTyxFQUFFdUIsa0JBQUEsQ0FBVUUsS0FBVixDQUFnQjtJQUN2QmQsTUFBTSxFQUFFWSxrQkFBQSxDQUFVRyxPQUFWLENBQ05ILGtCQUFBLENBQVVFLEtBQVYsQ0FBZ0I7TUFDZFIsR0FBRyxFQUFFTSxrQkFBQSxDQUFVQyxNQUFWLENBQWlCRyxVQURSO01BRWRYLEdBQUcsRUFBRU8sa0JBQUEsQ0FBVUssR0FGRDtNQUdkYixHQUFHLEVBQUVRLGtCQUFBLENBQVVDLE1BQVYsQ0FBaUJHO0lBSFIsQ0FBaEIsQ0FETSxDQURlO0lBUXZCWCxHQUFHLEVBQUVPLGtCQUFBLENBQVVLO0VBUlEsQ0FBaEIsRUFTTkQsVUFidUI7RUFjMUIxQixRQUFRLEVBQUVzQixrQkFBQSxDQUFVTTtBQWRNLENBQTVCO2VBaUJlL0IsZSJ9