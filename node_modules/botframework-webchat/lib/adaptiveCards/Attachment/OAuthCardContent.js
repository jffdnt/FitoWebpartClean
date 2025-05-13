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

var OAuthCardContent = function OAuthCardContent(_ref) {
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
      builder.addCommonHeaders(content);
      builder.addButtons(content.buttons, true);
      return builder.card;
    }
  }, [adaptiveCardsPackage, content, direction, styleOptions]);
  return /*#__PURE__*/_react.default.createElement(_AdaptiveCardRenderer.default, {
    actionPerformedClassName: actionPerformedClassName,
    adaptiveCard: builtCard,
    disabled: disabled
  });
};

OAuthCardContent.defaultProps = {
  actionPerformedClassName: '',
  disabled: undefined
};
OAuthCardContent.propTypes = {
  actionPerformedClassName: _propTypes.default.string,
  // PropTypes cannot fully capture TypeScript types.
  // @ts-ignore
  content: _propTypes.default.shape({
    buttons: _propTypes.default.array
  }).isRequired,
  disabled: _propTypes.default.bool
};
var _default = OAuthCardContent;
exports.default = _default;
//# sourceMappingURL=data:application/json;charset=utf-8;base64,eyJ2ZXJzaW9uIjozLCJuYW1lcyI6WyJ1c2VEaXJlY3Rpb24iLCJob29rcyIsIk9BdXRoQ2FyZENvbnRlbnQiLCJhY3Rpb25QZXJmb3JtZWRDbGFzc05hbWUiLCJjb250ZW50IiwiZGlzYWJsZWQiLCJ1c2VBZGFwdGl2ZUNhcmRzUGFja2FnZSIsImFkYXB0aXZlQ2FyZHNQYWNrYWdlIiwiZGlyZWN0aW9uIiwidXNlU3R5bGVPcHRpb25zIiwic3R5bGVPcHRpb25zIiwiYnVpbHRDYXJkIiwidXNlTWVtbyIsImJ1aWxkZXIiLCJBZGFwdGl2ZUNhcmRCdWlsZGVyIiwiYWRkQ29tbW9uSGVhZGVycyIsImFkZEJ1dHRvbnMiLCJidXR0b25zIiwiY2FyZCIsImRlZmF1bHRQcm9wcyIsInVuZGVmaW5lZCIsInByb3BUeXBlcyIsIlByb3BUeXBlcyIsInN0cmluZyIsInNoYXBlIiwiYXJyYXkiLCJpc1JlcXVpcmVkIiwiYm9vbCJdLCJzb3VyY2VSb290IjoiYnVuZGxlOi8vLyIsInNvdXJjZXMiOlsiLi4vLi4vLi4vc3JjL2FkYXB0aXZlQ2FyZHMvQXR0YWNobWVudC9PQXV0aENhcmRDb250ZW50LnRzeCJdLCJzb3VyY2VzQ29udGVudCI6WyJpbXBvcnQgeyBob29rcyB9IGZyb20gJ2JvdGZyYW1ld29yay13ZWJjaGF0LWNvbXBvbmVudCc7XG5pbXBvcnQgUHJvcFR5cGVzIGZyb20gJ3Byb3AtdHlwZXMnO1xuaW1wb3J0IFJlYWN0LCB7IEZDLCB1c2VNZW1vIH0gZnJvbSAncmVhY3QnO1xuaW1wb3J0IHR5cGUgeyBEaXJlY3RMaW5lT0F1dGhDYXJkIH0gZnJvbSAnYm90ZnJhbWV3b3JrLXdlYmNoYXQtY29yZSc7XG5cbmltcG9ydCBBZGFwdGl2ZUNhcmRCdWlsZGVyIGZyb20gJy4vQWRhcHRpdmVDYXJkQnVpbGRlcic7XG5pbXBvcnQgQWRhcHRpdmVDYXJkUmVuZGVyZXIgZnJvbSAnLi9BZGFwdGl2ZUNhcmRSZW5kZXJlcic7XG5pbXBvcnQgdXNlQWRhcHRpdmVDYXJkc1BhY2thZ2UgZnJvbSAnLi4vaG9va3MvdXNlQWRhcHRpdmVDYXJkc1BhY2thZ2UnO1xuaW1wb3J0IHVzZVN0eWxlT3B0aW9ucyBmcm9tICcuLi8uLi9ob29rcy91c2VTdHlsZU9wdGlvbnMnO1xuXG5jb25zdCB7IHVzZURpcmVjdGlvbiB9ID0gaG9va3M7XG5cbnR5cGUgT0F1dGhDYXJkQ29udGVudFByb3BzID0ge1xuICBhY3Rpb25QZXJmb3JtZWRDbGFzc05hbWU/OiBzdHJpbmc7XG4gIGNvbnRlbnQ6IERpcmVjdExpbmVPQXV0aENhcmQ7XG4gIGRpc2FibGVkPzogYm9vbGVhbjtcbn07XG5cbmNvbnN0IE9BdXRoQ2FyZENvbnRlbnQ6IEZDPE9BdXRoQ2FyZENvbnRlbnRQcm9wcz4gPSAoeyBhY3Rpb25QZXJmb3JtZWRDbGFzc05hbWUsIGNvbnRlbnQsIGRpc2FibGVkIH0pID0+IHtcbiAgY29uc3QgW2FkYXB0aXZlQ2FyZHNQYWNrYWdlXSA9IHVzZUFkYXB0aXZlQ2FyZHNQYWNrYWdlKCk7XG4gIGNvbnN0IFtkaXJlY3Rpb25dID0gdXNlRGlyZWN0aW9uKCk7XG4gIGNvbnN0IFtzdHlsZU9wdGlvbnNdID0gdXNlU3R5bGVPcHRpb25zKCk7XG5cbiAgY29uc3QgYnVpbHRDYXJkID0gdXNlTWVtbygoKSA9PiB7XG4gICAgaWYgKGNvbnRlbnQpIHtcbiAgICAgIGNvbnN0IGJ1aWxkZXIgPSBuZXcgQWRhcHRpdmVDYXJkQnVpbGRlcihhZGFwdGl2ZUNhcmRzUGFja2FnZSwgc3R5bGVPcHRpb25zLCBkaXJlY3Rpb24pO1xuXG4gICAgICBidWlsZGVyLmFkZENvbW1vbkhlYWRlcnMoY29udGVudCk7XG4gICAgICBidWlsZGVyLmFkZEJ1dHRvbnMoY29udGVudC5idXR0b25zLCB0cnVlKTtcblxuICAgICAgcmV0dXJuIGJ1aWxkZXIuY2FyZDtcbiAgICB9XG4gIH0sIFthZGFwdGl2ZUNhcmRzUGFja2FnZSwgY29udGVudCwgZGlyZWN0aW9uLCBzdHlsZU9wdGlvbnNdKTtcblxuICByZXR1cm4gKFxuICAgIDxBZGFwdGl2ZUNhcmRSZW5kZXJlclxuICAgICAgYWN0aW9uUGVyZm9ybWVkQ2xhc3NOYW1lPXthY3Rpb25QZXJmb3JtZWRDbGFzc05hbWV9XG4gICAgICBhZGFwdGl2ZUNhcmQ9e2J1aWx0Q2FyZH1cbiAgICAgIGRpc2FibGVkPXtkaXNhYmxlZH1cbiAgICAvPlxuICApO1xufTtcblxuT0F1dGhDYXJkQ29udGVudC5kZWZhdWx0UHJvcHMgPSB7XG4gIGFjdGlvblBlcmZvcm1lZENsYXNzTmFtZTogJycsXG4gIGRpc2FibGVkOiB1bmRlZmluZWRcbn07XG5cbk9BdXRoQ2FyZENvbnRlbnQucHJvcFR5cGVzID0ge1xuICBhY3Rpb25QZXJmb3JtZWRDbGFzc05hbWU6IFByb3BUeXBlcy5zdHJpbmcsXG4gIC8vIFByb3BUeXBlcyBjYW5ub3QgZnVsbHkgY2FwdHVyZSBUeXBlU2NyaXB0IHR5cGVzLlxuICAvLyBAdHMtaWdub3JlXG4gIGNvbnRlbnQ6IFByb3BUeXBlcy5zaGFwZSh7XG4gICAgYnV0dG9uczogUHJvcFR5cGVzLmFycmF5XG4gIH0pLmlzUmVxdWlyZWQsXG4gIGRpc2FibGVkOiBQcm9wVHlwZXMuYm9vbFxufTtcblxuZXhwb3J0IGRlZmF1bHQgT0F1dGhDYXJkQ29udGVudDtcbiJdLCJtYXBwaW5ncyI6Ijs7Ozs7Ozs7Ozs7OztBQUFBOztBQUNBOztBQUNBOztBQUdBOztBQUNBOztBQUNBOztBQUNBOzs7Ozs7QUFFQSxJQUFRQSxZQUFSLEdBQXlCQyxtQ0FBekIsQ0FBUUQsWUFBUjs7QUFRQSxJQUFNRSxnQkFBMkMsR0FBRyxTQUE5Q0EsZ0JBQThDLE9BQXFEO0VBQUEsSUFBbERDLHdCQUFrRCxRQUFsREEsd0JBQWtEO0VBQUEsSUFBeEJDLE9BQXdCLFFBQXhCQSxPQUF3QjtFQUFBLElBQWZDLFFBQWUsUUFBZkEsUUFBZTs7RUFDdkcsNEJBQStCLElBQUFDLGdDQUFBLEdBQS9CO0VBQUE7RUFBQSxJQUFPQyxvQkFBUDs7RUFDQSxvQkFBb0JQLFlBQVksRUFBaEM7RUFBQTtFQUFBLElBQU9RLFNBQVA7O0VBQ0EsdUJBQXVCLElBQUFDLHlCQUFBLEdBQXZCO0VBQUE7RUFBQSxJQUFPQyxZQUFQOztFQUVBLElBQU1DLFNBQVMsR0FBRyxJQUFBQyxjQUFBLEVBQVEsWUFBTTtJQUM5QixJQUFJUixPQUFKLEVBQWE7TUFDWCxJQUFNUyxPQUFPLEdBQUcsSUFBSUMsNEJBQUosQ0FBd0JQLG9CQUF4QixFQUE4Q0csWUFBOUMsRUFBNERGLFNBQTVELENBQWhCO01BRUFLLE9BQU8sQ0FBQ0UsZ0JBQVIsQ0FBeUJYLE9BQXpCO01BQ0FTLE9BQU8sQ0FBQ0csVUFBUixDQUFtQlosT0FBTyxDQUFDYSxPQUEzQixFQUFvQyxJQUFwQztNQUVBLE9BQU9KLE9BQU8sQ0FBQ0ssSUFBZjtJQUNEO0VBQ0YsQ0FUaUIsRUFTZixDQUFDWCxvQkFBRCxFQUF1QkgsT0FBdkIsRUFBZ0NJLFNBQWhDLEVBQTJDRSxZQUEzQyxDQVRlLENBQWxCO0VBV0Esb0JBQ0UsNkJBQUMsNkJBQUQ7SUFDRSx3QkFBd0IsRUFBRVAsd0JBRDVCO0lBRUUsWUFBWSxFQUFFUSxTQUZoQjtJQUdFLFFBQVEsRUFBRU47RUFIWixFQURGO0FBT0QsQ0F2QkQ7O0FBeUJBSCxnQkFBZ0IsQ0FBQ2lCLFlBQWpCLEdBQWdDO0VBQzlCaEIsd0JBQXdCLEVBQUUsRUFESTtFQUU5QkUsUUFBUSxFQUFFZTtBQUZvQixDQUFoQztBQUtBbEIsZ0JBQWdCLENBQUNtQixTQUFqQixHQUE2QjtFQUMzQmxCLHdCQUF3QixFQUFFbUIsa0JBQUEsQ0FBVUMsTUFEVDtFQUUzQjtFQUNBO0VBQ0FuQixPQUFPLEVBQUVrQixrQkFBQSxDQUFVRSxLQUFWLENBQWdCO0lBQ3ZCUCxPQUFPLEVBQUVLLGtCQUFBLENBQVVHO0VBREksQ0FBaEIsRUFFTkMsVUFOd0I7RUFPM0JyQixRQUFRLEVBQUVpQixrQkFBQSxDQUFVSztBQVBPLENBQTdCO2VBVWV6QixnQiJ9