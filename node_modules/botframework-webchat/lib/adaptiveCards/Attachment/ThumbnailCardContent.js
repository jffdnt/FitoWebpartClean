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

/* eslint no-magic-numbers: ["error", { "ignore": [25, 75] }] */
var useDirection = _botframeworkWebchatComponent.hooks.useDirection;

var ThumbnailCardContent = function ThumbnailCardContent(_ref) {
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
      var TextSize = adaptiveCardsPackage.TextSize,
          TextWeight = adaptiveCardsPackage.TextWeight;
      var buttons = content.buttons,
          images = content.images,
          subtitle = content.subtitle,
          text = content.text,
          title = content.title;
      var richCardWrapTitle = styleOptions.richCardWrapTitle;

      if (images && images.length) {
        var _builder$addColumnSet = builder.addColumnSet([75, 25]),
            _builder$addColumnSet2 = (0, _slicedToArray2.default)(_builder$addColumnSet, 2),
            firstColumn = _builder$addColumnSet2[0],
            lastColumn = _builder$addColumnSet2[1];

        var _images = (0, _slicedToArray2.default)(images, 1),
            _images$ = _images[0],
            alt = _images$.alt,
            tap = _images$.tap,
            url = _images$.url;

        builder.addTextBlock(title, {
          size: TextSize.Medium,
          weight: TextWeight.Bolder,
          wrap: richCardWrapTitle
        }, firstColumn);
        builder.addTextBlock(subtitle, {
          isSubtle: true,
          wrap: richCardWrapTitle
        }, firstColumn);
        builder.addImage(url, lastColumn, tap, alt);
        builder.addTextBlock(text, {
          wrap: true
        });
        builder.addButtons(buttons);
      } else {
        builder.addCommon(content);
      }

      return builder.card;
    }
  }, [adaptiveCardsPackage, direction, content, styleOptions]);
  return /*#__PURE__*/_react.default.createElement(_AdaptiveCardRenderer.default, {
    actionPerformedClassName: actionPerformedClassName,
    adaptiveCard: builtCard,
    disabled: disabled,
    tapAction: content && content.tap
  });
};

ThumbnailCardContent.defaultProps = {
  actionPerformedClassName: '',
  disabled: undefined
};
ThumbnailCardContent.propTypes = {
  actionPerformedClassName: _propTypes.default.string,
  // PropTypes cannot fully capture TypeScript types.
  // @ts-ignore
  content: _propTypes.default.shape({
    buttons: _propTypes.default.array,
    images: _propTypes.default.arrayOf(_propTypes.default.shape({
      alt: _propTypes.default.string.isRequired,
      tap: _propTypes.default.any,
      url: _propTypes.default.string.isRequired
    })),
    subtitle: _propTypes.default.string,
    tap: _propTypes.default.any,
    text: _propTypes.default.string,
    title: _propTypes.default.string
  }).isRequired,
  disabled: _propTypes.default.bool
};
var _default = ThumbnailCardContent;
exports.default = _default;
//# sourceMappingURL=data:application/json;charset=utf-8;base64,eyJ2ZXJzaW9uIjozLCJuYW1lcyI6WyJ1c2VEaXJlY3Rpb24iLCJob29rcyIsIlRodW1ibmFpbENhcmRDb250ZW50IiwiYWN0aW9uUGVyZm9ybWVkQ2xhc3NOYW1lIiwiY29udGVudCIsImRpc2FibGVkIiwidXNlQWRhcHRpdmVDYXJkc1BhY2thZ2UiLCJhZGFwdGl2ZUNhcmRzUGFja2FnZSIsImRpcmVjdGlvbiIsInVzZVN0eWxlT3B0aW9ucyIsInN0eWxlT3B0aW9ucyIsImJ1aWx0Q2FyZCIsInVzZU1lbW8iLCJidWlsZGVyIiwiQWRhcHRpdmVDYXJkQnVpbGRlciIsIlRleHRTaXplIiwiVGV4dFdlaWdodCIsImJ1dHRvbnMiLCJpbWFnZXMiLCJzdWJ0aXRsZSIsInRleHQiLCJ0aXRsZSIsInJpY2hDYXJkV3JhcFRpdGxlIiwibGVuZ3RoIiwiYWRkQ29sdW1uU2V0IiwiZmlyc3RDb2x1bW4iLCJsYXN0Q29sdW1uIiwiYWx0IiwidGFwIiwidXJsIiwiYWRkVGV4dEJsb2NrIiwic2l6ZSIsIk1lZGl1bSIsIndlaWdodCIsIkJvbGRlciIsIndyYXAiLCJpc1N1YnRsZSIsImFkZEltYWdlIiwiYWRkQnV0dG9ucyIsImFkZENvbW1vbiIsImNhcmQiLCJkZWZhdWx0UHJvcHMiLCJ1bmRlZmluZWQiLCJwcm9wVHlwZXMiLCJQcm9wVHlwZXMiLCJzdHJpbmciLCJzaGFwZSIsImFycmF5IiwiYXJyYXlPZiIsImlzUmVxdWlyZWQiLCJhbnkiLCJib29sIl0sInNvdXJjZVJvb3QiOiJidW5kbGU6Ly8vIiwic291cmNlcyI6WyIuLi8uLi8uLi9zcmMvYWRhcHRpdmVDYXJkcy9BdHRhY2htZW50L1RodW1ibmFpbENhcmRDb250ZW50LnRzeCJdLCJzb3VyY2VzQ29udGVudCI6WyIvKiBlc2xpbnQgbm8tbWFnaWMtbnVtYmVyczogW1wiZXJyb3JcIiwgeyBcImlnbm9yZVwiOiBbMjUsIDc1XSB9XSAqL1xuXG5pbXBvcnQgeyBob29rcyB9IGZyb20gJ2JvdGZyYW1ld29yay13ZWJjaGF0LWNvbXBvbmVudCc7XG5pbXBvcnQgUHJvcFR5cGVzIGZyb20gJ3Byb3AtdHlwZXMnO1xuaW1wb3J0IFJlYWN0LCB7IEZDLCB1c2VNZW1vIH0gZnJvbSAncmVhY3QnO1xuaW1wb3J0IHR5cGUgeyBEaXJlY3RMaW5lVGh1bWJuYWlsQ2FyZCB9IGZyb20gJ2JvdGZyYW1ld29yay13ZWJjaGF0LWNvcmUnO1xuXG5pbXBvcnQgQWRhcHRpdmVDYXJkQnVpbGRlciBmcm9tICcuL0FkYXB0aXZlQ2FyZEJ1aWxkZXInO1xuaW1wb3J0IEFkYXB0aXZlQ2FyZFJlbmRlcmVyIGZyb20gJy4vQWRhcHRpdmVDYXJkUmVuZGVyZXInO1xuaW1wb3J0IHVzZUFkYXB0aXZlQ2FyZHNQYWNrYWdlIGZyb20gJy4uL2hvb2tzL3VzZUFkYXB0aXZlQ2FyZHNQYWNrYWdlJztcbmltcG9ydCB1c2VTdHlsZU9wdGlvbnMgZnJvbSAnLi4vLi4vaG9va3MvdXNlU3R5bGVPcHRpb25zJztcblxuY29uc3QgeyB1c2VEaXJlY3Rpb24gfSA9IGhvb2tzO1xuXG50eXBlIFRodW1ibmFpbENhcmRDb250ZW50UHJvcHMgPSB7XG4gIGFjdGlvblBlcmZvcm1lZENsYXNzTmFtZT86IHN0cmluZztcbiAgY29udGVudDogRGlyZWN0TGluZVRodW1ibmFpbENhcmQ7XG4gIGRpc2FibGVkPzogYm9vbGVhbjtcbn07XG5cbmNvbnN0IFRodW1ibmFpbENhcmRDb250ZW50OiBGQzxUaHVtYm5haWxDYXJkQ29udGVudFByb3BzPiA9ICh7IGFjdGlvblBlcmZvcm1lZENsYXNzTmFtZSwgY29udGVudCwgZGlzYWJsZWQgfSkgPT4ge1xuICBjb25zdCBbYWRhcHRpdmVDYXJkc1BhY2thZ2VdID0gdXNlQWRhcHRpdmVDYXJkc1BhY2thZ2UoKTtcbiAgY29uc3QgW2RpcmVjdGlvbl0gPSB1c2VEaXJlY3Rpb24oKTtcbiAgY29uc3QgW3N0eWxlT3B0aW9uc10gPSB1c2VTdHlsZU9wdGlvbnMoKTtcblxuICBjb25zdCBidWlsdENhcmQgPSB1c2VNZW1vKCgpID0+IHtcbiAgICBpZiAoY29udGVudCkge1xuICAgICAgY29uc3QgYnVpbGRlciA9IG5ldyBBZGFwdGl2ZUNhcmRCdWlsZGVyKGFkYXB0aXZlQ2FyZHNQYWNrYWdlLCBzdHlsZU9wdGlvbnMsIGRpcmVjdGlvbik7XG4gICAgICBjb25zdCB7IFRleHRTaXplLCBUZXh0V2VpZ2h0IH0gPSBhZGFwdGl2ZUNhcmRzUGFja2FnZTtcbiAgICAgIGNvbnN0IHsgYnV0dG9ucywgaW1hZ2VzLCBzdWJ0aXRsZSwgdGV4dCwgdGl0bGUgfSA9IGNvbnRlbnQ7XG4gICAgICBjb25zdCB7IHJpY2hDYXJkV3JhcFRpdGxlIH0gPSBzdHlsZU9wdGlvbnM7XG5cbiAgICAgIGlmIChpbWFnZXMgJiYgaW1hZ2VzLmxlbmd0aCkge1xuICAgICAgICBjb25zdCBbZmlyc3RDb2x1bW4sIGxhc3RDb2x1bW5dID0gYnVpbGRlci5hZGRDb2x1bW5TZXQoWzc1LCAyNV0pO1xuICAgICAgICBjb25zdCBbeyBhbHQsIHRhcCwgdXJsIH1dID0gaW1hZ2VzO1xuXG4gICAgICAgIGJ1aWxkZXIuYWRkVGV4dEJsb2NrKFxuICAgICAgICAgIHRpdGxlLFxuICAgICAgICAgIHsgc2l6ZTogVGV4dFNpemUuTWVkaXVtLCB3ZWlnaHQ6IFRleHRXZWlnaHQuQm9sZGVyLCB3cmFwOiByaWNoQ2FyZFdyYXBUaXRsZSB9LFxuICAgICAgICAgIGZpcnN0Q29sdW1uXG4gICAgICAgICk7XG5cbiAgICAgICAgYnVpbGRlci5hZGRUZXh0QmxvY2soc3VidGl0bGUsIHsgaXNTdWJ0bGU6IHRydWUsIHdyYXA6IHJpY2hDYXJkV3JhcFRpdGxlIH0sIGZpcnN0Q29sdW1uKTtcbiAgICAgICAgYnVpbGRlci5hZGRJbWFnZSh1cmwsIGxhc3RDb2x1bW4sIHRhcCwgYWx0KTtcbiAgICAgICAgYnVpbGRlci5hZGRUZXh0QmxvY2sodGV4dCwgeyB3cmFwOiB0cnVlIH0pO1xuICAgICAgICBidWlsZGVyLmFkZEJ1dHRvbnMoYnV0dG9ucyk7XG4gICAgICB9IGVsc2Uge1xuICAgICAgICBidWlsZGVyLmFkZENvbW1vbihjb250ZW50KTtcbiAgICAgIH1cbiAgICAgIHJldHVybiBidWlsZGVyLmNhcmQ7XG4gICAgfVxuICB9LCBbYWRhcHRpdmVDYXJkc1BhY2thZ2UsIGRpcmVjdGlvbiwgY29udGVudCwgc3R5bGVPcHRpb25zXSk7XG5cbiAgcmV0dXJuIChcbiAgICA8QWRhcHRpdmVDYXJkUmVuZGVyZXJcbiAgICAgIGFjdGlvblBlcmZvcm1lZENsYXNzTmFtZT17YWN0aW9uUGVyZm9ybWVkQ2xhc3NOYW1lfVxuICAgICAgYWRhcHRpdmVDYXJkPXtidWlsdENhcmR9XG4gICAgICBkaXNhYmxlZD17ZGlzYWJsZWR9XG4gICAgICB0YXBBY3Rpb249e2NvbnRlbnQgJiYgY29udGVudC50YXB9XG4gICAgLz5cbiAgKTtcbn07XG5cblRodW1ibmFpbENhcmRDb250ZW50LmRlZmF1bHRQcm9wcyA9IHtcbiAgYWN0aW9uUGVyZm9ybWVkQ2xhc3NOYW1lOiAnJyxcbiAgZGlzYWJsZWQ6IHVuZGVmaW5lZFxufTtcblxuVGh1bWJuYWlsQ2FyZENvbnRlbnQucHJvcFR5cGVzID0ge1xuICBhY3Rpb25QZXJmb3JtZWRDbGFzc05hbWU6IFByb3BUeXBlcy5zdHJpbmcsXG4gIC8vIFByb3BUeXBlcyBjYW5ub3QgZnVsbHkgY2FwdHVyZSBUeXBlU2NyaXB0IHR5cGVzLlxuICAvLyBAdHMtaWdub3JlXG4gIGNvbnRlbnQ6IFByb3BUeXBlcy5zaGFwZSh7XG4gICAgYnV0dG9uczogUHJvcFR5cGVzLmFycmF5LFxuICAgIGltYWdlczogUHJvcFR5cGVzLmFycmF5T2YoXG4gICAgICBQcm9wVHlwZXMuc2hhcGUoe1xuICAgICAgICBhbHQ6IFByb3BUeXBlcy5zdHJpbmcuaXNSZXF1aXJlZCxcbiAgICAgICAgdGFwOiBQcm9wVHlwZXMuYW55LFxuICAgICAgICB1cmw6IFByb3BUeXBlcy5zdHJpbmcuaXNSZXF1aXJlZFxuICAgICAgfSlcbiAgICApLFxuICAgIHN1YnRpdGxlOiBQcm9wVHlwZXMuc3RyaW5nLFxuICAgIHRhcDogUHJvcFR5cGVzLmFueSxcbiAgICB0ZXh0OiBQcm9wVHlwZXMuc3RyaW5nLFxuICAgIHRpdGxlOiBQcm9wVHlwZXMuc3RyaW5nXG4gIH0pLmlzUmVxdWlyZWQsXG4gIGRpc2FibGVkOiBQcm9wVHlwZXMuYm9vbFxufTtcblxuZXhwb3J0IGRlZmF1bHQgVGh1bWJuYWlsQ2FyZENvbnRlbnQ7XG4iXSwibWFwcGluZ3MiOiI7Ozs7Ozs7Ozs7Ozs7QUFFQTs7QUFDQTs7QUFDQTs7QUFHQTs7QUFDQTs7QUFDQTs7QUFDQTs7Ozs7O0FBVkE7QUFZQSxJQUFRQSxZQUFSLEdBQXlCQyxtQ0FBekIsQ0FBUUQsWUFBUjs7QUFRQSxJQUFNRSxvQkFBbUQsR0FBRyxTQUF0REEsb0JBQXNELE9BQXFEO0VBQUEsSUFBbERDLHdCQUFrRCxRQUFsREEsd0JBQWtEO0VBQUEsSUFBeEJDLE9BQXdCLFFBQXhCQSxPQUF3QjtFQUFBLElBQWZDLFFBQWUsUUFBZkEsUUFBZTs7RUFDL0csNEJBQStCLElBQUFDLGdDQUFBLEdBQS9CO0VBQUE7RUFBQSxJQUFPQyxvQkFBUDs7RUFDQSxvQkFBb0JQLFlBQVksRUFBaEM7RUFBQTtFQUFBLElBQU9RLFNBQVA7O0VBQ0EsdUJBQXVCLElBQUFDLHlCQUFBLEdBQXZCO0VBQUE7RUFBQSxJQUFPQyxZQUFQOztFQUVBLElBQU1DLFNBQVMsR0FBRyxJQUFBQyxjQUFBLEVBQVEsWUFBTTtJQUM5QixJQUFJUixPQUFKLEVBQWE7TUFDWCxJQUFNUyxPQUFPLEdBQUcsSUFBSUMsNEJBQUosQ0FBd0JQLG9CQUF4QixFQUE4Q0csWUFBOUMsRUFBNERGLFNBQTVELENBQWhCO01BQ0EsSUFBUU8sUUFBUixHQUFpQ1Isb0JBQWpDLENBQVFRLFFBQVI7TUFBQSxJQUFrQkMsVUFBbEIsR0FBaUNULG9CQUFqQyxDQUFrQlMsVUFBbEI7TUFDQSxJQUFRQyxPQUFSLEdBQW1EYixPQUFuRCxDQUFRYSxPQUFSO01BQUEsSUFBaUJDLE1BQWpCLEdBQW1EZCxPQUFuRCxDQUFpQmMsTUFBakI7TUFBQSxJQUF5QkMsUUFBekIsR0FBbURm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