"use strict";

var _interopRequireDefault = require("@babel/runtime/helpers/interopRequireDefault");

var _typeof3 = require("@babel/runtime/helpers/typeof");

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.default = void 0;

var _typeof2 = _interopRequireDefault(require("@babel/runtime/helpers/typeof"));

var _defineProperty2 = _interopRequireDefault(require("@babel/runtime/helpers/defineProperty"));

var _propTypes = _interopRequireDefault(require("prop-types"));

var _react = _interopRequireWildcard(require("react"));

var _AdaptiveCardRenderer = _interopRequireDefault(require("./AdaptiveCardRenderer"));

var _useParseAdaptiveCardJSON = _interopRequireDefault(require("../hooks/internal/useParseAdaptiveCardJSON"));

function _getRequireWildcardCache(nodeInterop) { if (typeof WeakMap !== "function") return null; var cacheBabelInterop = new WeakMap(); var cacheNodeInterop = new WeakMap(); return (_getRequireWildcardCache = function _getRequireWildcardCache(nodeInterop) { return nodeInterop ? cacheNodeInterop : cacheBabelInterop; })(nodeInterop); }

function _interopRequireWildcard(obj, nodeInterop) { if (!nodeInterop && obj && obj.__esModule) { return obj; } if (obj === null || _typeof3(obj) !== "object" && typeof obj !== "function") { return { default: obj }; } var cache = _getRequireWildcardCache(nodeInterop); if (cache && cache.has(obj)) { return cache.get(obj); } var newObj = {}; var hasPropertyDescriptor = Object.defineProperty && Object.getOwnPropertyDescriptor; for (var key in obj) { if (key !== "default" && Object.prototype.hasOwnProperty.call(obj, key)) { var desc = hasPropertyDescriptor ? Object.getOwnPropertyDescriptor(obj, key) : null; if (desc && (desc.get || desc.set)) { Object.defineProperty(newObj, key, desc); } else { newObj[key] = obj[key]; } } } newObj.default = obj; if (cache) { cache.set(obj, newObj); } return newObj; }

function ownKeys(object, enumerableOnly) { var keys = Object.keys(object); if (Object.getOwnPropertySymbols) { var symbols = Object.getOwnPropertySymbols(object); enumerableOnly && (symbols = symbols.filter(function (sym) { return Object.getOwnPropertyDescriptor(object, sym).enumerable; })), keys.push.apply(keys, symbols); } return keys; }

function _objectSpread(target) { for (var i = 1; i < arguments.length; i++) { var source = null != arguments[i] ? arguments[i] : {}; i % 2 ? ownKeys(Object(source), !0).forEach(function (key) { (0, _defineProperty2.default)(target, key, source[key]); }) : Object.getOwnPropertyDescriptors ? Object.defineProperties(target, Object.getOwnPropertyDescriptors(source)) : ownKeys(Object(source)).forEach(function (key) { Object.defineProperty(target, key, Object.getOwnPropertyDescriptor(source, key)); }); } return target; }

function stripSubmitAction(card) {
  if (!card.actions) {
    return card;
  } // Filter out HTTP action buttons


  var nextActions = card.actions.filter(function (action) {
    return action.type !== 'Action.Submit';
  }).map(function (action) {
    return action.type === 'Action.ShowCard' ? _objectSpread(_objectSpread({}, action), {}, {
      card: stripSubmitAction(action.card)
    }) : action;
  });
  return _objectSpread(_objectSpread({}, card), {}, {
    nextActions: nextActions
  });
}

var AdaptiveCardContent = function AdaptiveCardContent(_ref) {
  var actionPerformedClassName = _ref.actionPerformedClassName,
      content = _ref.content,
      disabled = _ref.disabled;
  var parseAdaptiveCardJSON = (0, _useParseAdaptiveCardJSON.default)();
  var card = (0, _react.useMemo)(function () {
    return parseAdaptiveCardJSON(stripSubmitAction(_objectSpread({
      version: '1.0'
    }, (0, _typeof2.default)(content) === 'object' ? content : {})), {
      ignoreErrors: true
    });
  }, [content, parseAdaptiveCardJSON]);
  return !!card && /*#__PURE__*/_react.default.createElement(_AdaptiveCardRenderer.default, {
    actionPerformedClassName: actionPerformedClassName,
    adaptiveCard: card,
    disabled: disabled
  });
};

AdaptiveCardContent.defaultProps = {
  actionPerformedClassName: '',
  disabled: undefined
};
AdaptiveCardContent.propTypes = {
  actionPerformedClassName: _propTypes.default.string,
  content: _propTypes.default.any.isRequired,
  disabled: _propTypes.default.bool
};
var _default = AdaptiveCardContent;
exports.default = _default;
//# sourceMappingURL=data:application/json;charset=utf-8;base64,eyJ2ZXJzaW9uIjozLCJuYW1lcyI6WyJzdHJpcFN1Ym1pdEFjdGlvbiIsImNhcmQiLCJhY3Rpb25zIiwibmV4dEFjdGlvbnMiLCJmaWx0ZXIiLCJhY3Rpb24iLCJ0eXBlIiwibWFwIiwiQWRhcHRpdmVDYXJkQ29udGVudCIsImFjdGlvblBlcmZvcm1lZENsYXNzTmFtZSIsImNvbnRlbnQiLCJkaXNhYmxlZCIsInBhcnNlQWRhcHRpdmVDYXJkSlNPTiIsInVzZVBhcnNlQWRhcHRpdmVDYXJkSlNPTiIsInVzZU1lbW8iLCJ2ZXJzaW9uIiwiaWdub3JlRXJyb3JzIiwiZGVmYXVsdFByb3BzIiwidW5kZWZpbmVkIiwicHJvcFR5cGVzIiwiUHJvcFR5cGVzIiwic3RyaW5nIiwiYW55IiwiaXNSZXF1aXJlZCIsImJvb2wiXSwic291cmNlUm9vdCI6ImJ1bmRsZTovLy8iLCJzb3VyY2VzIjpbIi4uLy4uLy4uL3NyYy9hZGFwdGl2ZUNhcmRzL0F0dGFjaG1lbnQvQWRhcHRpdmVDYXJkQ29udGVudC50c3giXSwic291cmNlc0NvbnRlbnQiOlsiaW1wb3J0IFByb3BUeXBlcyBmcm9tICdwcm9wLXR5cGVzJztcbmltcG9ydCBSZWFjdCwgeyBGQywgdXNlTWVtbyB9IGZyb20gJ3JlYWN0JztcblxuaW1wb3J0IEFkYXB0aXZlQ2FyZFJlbmRlcmVyIGZyb20gJy4vQWRhcHRpdmVDYXJkUmVuZGVyZXInO1xuaW1wb3J0IHVzZVBhcnNlQWRhcHRpdmVDYXJkSlNPTiBmcm9tICcuLi9ob29rcy9pbnRlcm5hbC91c2VQYXJzZUFkYXB0aXZlQ2FyZEpTT04nO1xuXG5mdW5jdGlvbiBzdHJpcFN1Ym1pdEFjdGlvbihjYXJkKSB7XG4gIGlmICghY2FyZC5hY3Rpb25zKSB7XG4gICAgcmV0dXJuIGNhcmQ7XG4gIH1cblxuICAvLyBGaWx0ZXIgb3V0IEhUVFAgYWN0aW9uIGJ1dHRvbnNcbiAgY29uc3QgbmV4dEFjdGlvbnMgPSBjYXJkLmFjdGlvbnNcbiAgICAuZmlsdGVyKGFjdGlvbiA9PiBhY3Rpb24udHlwZSAhPT0gJ0FjdGlvbi5TdWJtaXQnKVxuICAgIC5tYXAoYWN0aW9uID0+IChhY3Rpb24udHlwZSA9PT0gJ0FjdGlvbi5TaG93Q2FyZCcgPyB7IC4uLmFjdGlvbiwgY2FyZDogc3RyaXBTdWJtaXRBY3Rpb24oYWN0aW9uLmNhcmQpIH0gOiBhY3Rpb24pKTtcblxuICByZXR1cm4geyAuLi5jYXJkLCBuZXh0QWN0aW9ucyB9O1xufVxuXG50eXBlIEFkYXB0aXZlQ2FyZENvbnRlbnRQcm9wcyA9IHtcbiAgYWN0aW9uUGVyZm9ybWVkQ2xhc3NOYW1lPzogc3RyaW5nO1xuICBjb250ZW50OiBhbnk7XG4gIGRpc2FibGVkPzogYm9vbGVhbjtcbn07XG5cbmNvbnN0IEFkYXB0aXZlQ2FyZENvbnRlbnQ6IEZDPEFkYXB0aXZlQ2FyZENvbnRlbnRQcm9wcz4gPSAoeyBhY3Rpb25QZXJmb3JtZWRDbGFzc05hbWUsIGNvbnRlbnQsIGRpc2FibGVkIH0pID0+IHtcbiAgY29uc3QgcGFyc2VBZGFwdGl2ZUNhcmRKU09OID0gdXNlUGFyc2VBZGFwdGl2ZUNhcmRKU09OKCk7XG5cbiAgY29uc3QgY2FyZCA9IHVzZU1lbW8oXG4gICAgKCkgPT5cbiAgICAgIHBhcnNlQWRhcHRpdmVDYXJkSlNPTihcbiAgICAgICAgc3RyaXBTdWJtaXRBY3Rpb24oe1xuICAgICAgICAgIHZlcnNpb246ICcxLjAnLFxuICAgICAgICAgIC4uLih0eXBlb2YgY29udGVudCA9PT0gJ29iamVjdCcgPyBjb250ZW50IDoge30pXG4gICAgICAgIH0pLFxuICAgICAgICB7IGlnbm9yZUVycm9yczogdHJ1ZSB9XG4gICAgICApLFxuICAgIFtjb250ZW50LCBwYXJzZUFkYXB0aXZlQ2FyZEpTT05dXG4gICk7XG5cbiAgcmV0dXJuIChcbiAgICAhIWNhcmQgJiYgKFxuICAgICAgPEFkYXB0aXZlQ2FyZFJlbmRlcmVyXG4gICAgICAgIGFjdGlvblBlcmZvcm1lZENsYXNzTmFtZT17YWN0aW9uUGVyZm9ybWVkQ2xhc3NOYW1lfVxuICAgICAgICBhZGFwdGl2ZUNhcmQ9e2NhcmR9XG4gICAgICAgIGRpc2FibGVkPXtkaXNhYmxlZH1cbiAgICAgIC8+XG4gICAgKVxuICApO1xufTtcblxuQWRhcHRpdmVDYXJkQ29udGVudC5kZWZhdWx0UHJvcHMgPSB7XG4gIGFjdGlvblBlcmZvcm1lZENsYXNzTmFtZTogJycsXG4gIGRpc2FibGVkOiB1bmRlZmluZWRcbn07XG5cbkFkYXB0aXZlQ2FyZENvbnRlbnQucHJvcFR5cGVzID0ge1xuICBhY3Rpb25QZXJmb3JtZWRDbGFzc05hbWU6IFByb3BUeXBlcy5zdHJpbmcsXG4gIGNvbnRlbnQ6IFByb3BUeXBlcy5hbnkuaXNSZXF1aXJlZCxcbiAgZGlzYWJsZWQ6IFByb3BUeXBlcy5ib29sXG59O1xuXG5leHBvcnQgZGVmYXVsdCBBZGFwdGl2ZUNhcmRDb250ZW50O1xuIl0sIm1hcHBpbmdzIjoiOzs7Ozs7Ozs7Ozs7Ozs7QUFBQTs7QUFDQTs7QUFFQTs7QUFDQTs7Ozs7Ozs7OztBQUVBLFNBQVNBLGlCQUFULENBQTJCQyxJQUEzQixFQUFpQztFQUMvQixJQUFJLENBQUNBLElBQUksQ0FBQ0MsT0FBVixFQUFtQjtJQUNqQixPQUFPRCxJQUFQO0VBQ0QsQ0FIOEIsQ0FLL0I7OztFQUNBLElBQU1FLFdBQVcsR0FBR0YsSUFBSSxDQUFDQyxPQUFMLENBQ2pCRSxNQURpQixDQUNWLFVBQUFDLE1BQU07SUFBQSxPQUFJQSxNQUFNLENBQUNDLElBQVAsS0FBZ0IsZUFBcEI7RUFBQSxDQURJLEVBRWpCQyxHQUZpQixDQUViLFVBQUFGLE1BQU07SUFBQSxPQUFLQSxNQUFNLENBQUNDLElBQVAsS0FBZ0IsaUJBQWhCLG1DQUF5Q0QsTUFBekM7TUFBaURKLElBQUksRUFBRUQsaUJBQWlCLENBQUNLLE1BQU0sQ0FBQ0osSUFBUjtJQUF4RSxLQUEwRkksTUFBL0Y7RUFBQSxDQUZPLENBQXBCO0VBSUEsdUNBQVlKLElBQVo7SUFBa0JFLFdBQVcsRUFBWEE7RUFBbEI7QUFDRDs7QUFRRCxJQUFNSyxtQkFBaUQsR0FBRyxTQUFwREEsbUJBQW9ELE9BQXFEO0VBQUEsSUFBbERDLHdCQUFrRCxRQUFsREEsd0JBQWtEO0VBQUEsSUFBeEJDLE9BQXdCLFFBQXhCQSxPQUF3QjtFQUFBLElBQWZDLFFBQWUsUUFBZkEsUUFBZTtFQUM3RyxJQUFNQyxxQkFBcUIsR0FBRyxJQUFBQyxpQ0FBQSxHQUE5QjtFQUVBLElBQU1aLElBQUksR0FBRyxJQUFBYSxjQUFBLEVBQ1g7SUFBQSxPQUNFRixxQkFBcUIsQ0FDbkJaLGlCQUFpQjtNQUNmZSxPQUFPLEVBQUU7SUFETSxHQUVYLHNCQUFPTCxPQUFQLE1BQW1CLFFBQW5CLEdBQThCQSxPQUE5QixHQUF3QyxFQUY3QixFQURFLEVBS25CO01BQUVNLFlBQVksRUFBRTtJQUFoQixDQUxtQixDQUR2QjtFQUFBLENBRFcsRUFTWCxDQUFDTixPQUFELEVBQVVFLHFCQUFWLENBVFcsQ0FBYjtFQVlBLE9BQ0UsQ0FBQyxDQUFDWCxJQUFGLGlCQUNFLDZCQUFDLDZCQUFEO0lBQ0Usd0JBQXdCLEVBQUVRLHdCQUQ1QjtJQUVFLFlBQVksRUFBRVIsSUFGaEI7SUFHRSxRQUFRLEVBQUVVO0VBSFosRUFGSjtBQVNELENBeEJEOztBQTBCQUgsbUJBQW1CLENBQUNTLFlBQXBCLEdBQW1DO0VBQ2pDUix3QkFBd0IsRUFBRSxFQURPO0VBRWpDRSxRQUFRLEVBQUVPO0FBRnVCLENBQW5DO0FBS0FWLG1CQUFtQixDQUFDVyxTQUFwQixHQUFnQztFQUM5QlYsd0JBQXdCLEVBQUVXLGtCQUFBLENBQVVDLE1BRE47RUFFOUJYLE9BQU8sRUFBRVUsa0JBQUEsQ0FBVUUsR0FBVixDQUFjQyxVQUZPO0VBRzlCWixRQUFRLEVBQUVTLGtCQUFBLENBQVVJO0FBSFUsQ0FBaEM7ZUFNZWhCLG1CIn0=