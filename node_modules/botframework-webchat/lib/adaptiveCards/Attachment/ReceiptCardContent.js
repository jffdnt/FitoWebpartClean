"use strict";

var _interopRequireDefault = require("@babel/runtime/helpers/interopRequireDefault");

var _typeof = require("@babel/runtime/helpers/typeof");

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.default = void 0;

var _toArray2 = _interopRequireDefault(require("@babel/runtime/helpers/toArray"));

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

/* eslint no-magic-numbers: ["error", { "ignore": [0, 1, 10, 15, 25, 50, 75] }] */
var useDirection = _botframeworkWebchatComponent.hooks.useDirection,
    useLocalizer = _botframeworkWebchatComponent.hooks.useLocalizer;

function nullOrUndefined(obj) {
  return obj === null || typeof obj === 'undefined';
}

var ReceiptCardContent = function ReceiptCardContent(_ref) {
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

  var localize = useLocalizer();
  var taxText = localize('RECEIPT_CARD_TAX');
  var totalText = localize('RECEIPT_CARD_TOTAL');
  var vatText = localize('RECEIPT_CARD_VAT');
  var builtCard = (0, _react.useMemo)(function () {
    var builder = new _AdaptiveCardBuilder.default(adaptiveCardsPackage, styleOptions, direction);
    var HorizontalAlignment = adaptiveCardsPackage.HorizontalAlignment,
        TextSize = adaptiveCardsPackage.TextSize,
        TextWeight = adaptiveCardsPackage.TextWeight;
    var buttons = content.buttons,
        facts = content.facts,
        items = content.items,
        tax = content.tax,
        title = content.title,
        total = content.total,
        vat = content.vat;
    var richCardWrapTitle = styleOptions.richCardWrapTitle;

    if (content) {
      builder.addTextBlock(title, {
        size: TextSize.Medium,
        weight: TextWeight.Bolder,
        wrap: richCardWrapTitle
      });

      if (facts) {
        var _builder$addColumnSet = builder.addColumnSet([75, 25]),
            _builder$addColumnSet2 = (0, _slicedToArray2.default)(_builder$addColumnSet, 2),
            firstFactColumn = _builder$addColumnSet2[0],
            lastFactColumn = _builder$addColumnSet2[1];

        facts.map(function (_ref2) {
          var key = _ref2.key,
              value = _ref2.value;
          builder.addTextBlock(key, {
            size: TextSize.Medium
          }, firstFactColumn);
          builder.addTextBlock(value, {
            size: TextSize.Medium,
            horizontalAlignment: HorizontalAlignment.Right
          }, lastFactColumn);
        });
      }

      items && items.map(function (_ref3) {
        var _ref3$image = _ref3.image;
        _ref3$image = _ref3$image === void 0 ? {} : _ref3$image;
        var alt = _ref3$image.alt,
            imageTap = _ref3$image.tap,
            url = _ref3$image.url,
            price = _ref3.price,
            quantity = _ref3.quantity,
            subtitle = _ref3.subtitle,
            tap = _ref3.tap,
            text = _ref3.text,
            title = _ref3.title;
        var itemColumns;

        if (url) {
          var _builder$addColumnSet3 = builder.addColumnSet([15, 75, 10]),
              _builder$addColumnSet4 = (0, _toArray2.default)(_builder$addColumnSet3),
              itemImageColumn = _builder$addColumnSet4[0],
              columns = _builder$addColumnSet4.slice(1);

          itemColumns = columns;
          builder.addImage(url, itemImageColumn, imageTap, alt);
        } else {
          itemColumns = builder.addColumnSet([75, 25], undefined, tap && tap);
        }

        var _itemColumns = itemColumns,
            _itemColumns2 = (0, _slicedToArray2.default)(_itemColumns, 2),
            itemTitleColumn = _itemColumns2[0],
            itemPriceColumn = _itemColumns2[1];

        builder.addTextBlock(quantity ? "".concat(title, " &times; ").concat(quantity) : title, {
          size: TextSize.Medium,
          weight: TextWeight.Bolder,
          wrap: richCardWrapTitle
        }, itemTitleColumn);
        builder.addTextBlock(subtitle, {
          size: TextSize.Medium,
          wrap: richCardWrapTitle
        }, itemTitleColumn);
        builder.addTextBlock(price, {
          horizontalAlignment: HorizontalAlignment.Right
        }, itemPriceColumn);

        if (text) {
          builder.addTextBlock(text, {
            size: TextSize.Medium,
            wrap: richCardWrapTitle
          }, itemTitleColumn);
        }
      });

      if (!nullOrUndefined(vat)) {
        var vatCol = builder.addColumnSet([75, 25]);
        builder.addTextBlock(vatText, {
          size: TextSize.Medium,
          weight: TextWeight.Bolder
        }, vatCol[0]);
        builder.addTextBlock(vat, {
          horizontalAlignment: HorizontalAlignment.Right
        }, vatCol[1]);
      }

      if (!nullOrUndefined(tax)) {
        var taxCol = builder.addColumnSet([75, 25]);
        builder.addTextBlock(taxText, {
          size: TextSize.Medium,
          weight: TextWeight.Bolder
        }, taxCol[0]);
        builder.addTextBlock(tax, {
          horizontalAlignment: HorizontalAlignment.Right
        }, taxCol[1]);
      }

      if (!nullOrUndefined(total)) {
        var totalCol = builder.addColumnSet([75, 25]);
        builder.addTextBlock(totalText, {
          size: TextSize.Medium,
          weight: TextWeight.Bolder
        }, totalCol[0]);
        builder.addTextBlock(total, {
          horizontalAlignment: HorizontalAlignment.Right,
          size: TextSize.Medium,
          weight: TextWeight.Bolder
        }, totalCol[1]);
      }

      builder.addButtons(buttons);
      return builder.card;
    }
  }, [adaptiveCardsPackage, content, direction, styleOptions, taxText, totalText, vatText]);
  return /*#__PURE__*/_react.default.createElement(_AdaptiveCardRenderer.default, {
    actionPerformedClassName: actionPerformedClassName,
    adaptiveCard: builtCard,
    disabled: disabled,
    tapAction: content && content.tap
  });
};

ReceiptCardContent.defaultProps = {
  actionPerformedClassName: '',
  disabled: undefined
};
ReceiptCardContent.propTypes = {
  actionPerformedClassName: _propTypes.default.string,
  // PropTypes cannot fully capture TypeScript types.
  // @ts-ignore
  content: _propTypes.default.shape({
    buttons: _propTypes.default.array,
    facts: _propTypes.default.arrayOf(_propTypes.default.shape({
      key: _propTypes.default.string,
      value: _propTypes.default.string
    })),
    items: _propTypes.default.arrayOf(_propTypes.default.shape({
      image: _propTypes.default.shape({
        alt: _propTypes.default.string.isRequired,
        tap: _propTypes.default.any,
        url: _propTypes.default.string.isRequired
      }),
      price: _propTypes.default.string.isRequired,
      quantity: _propTypes.default.string,
      subtitle: _propTypes.default.string,
      tap: _propTypes.default.any,
      text: _propTypes.default.string,
      title: _propTypes.default.string.isRequired
    })),
    tap: _propTypes.default.any,
    tax: _propTypes.default.string,
    title: _propTypes.default.string,
    total: _propTypes.default.string,
    vat: _propTypes.default.string
  }).isRequired,
  disabled: _propTypes.default.bool
};
var _default = ReceiptCardContent;
exports.default = _default;
//# sourceMappingURL=data:application/json;charset=utf-8;base64,eyJ2ZXJzaW9uIjozLCJuYW1lcyI6WyJ1c2VEaXJlY3Rpb24iLCJob29rcyIsInVzZUxvY2FsaXplciIsIm51bGxPclVuZGVmaW5lZCIsIm9iaiIsIlJlY2VpcHRDYXJkQ29udGVudCIsImFjdGlvblBlcmZvcm1lZENsYXNzTmFtZSIsImNvbnRlbnQiLCJkaXNhYmxlZCIsInVzZUFkYXB0aXZlQ2FyZHNQYWNrYWdlIiwiYWRhcHRpdmVDYXJkc1BhY2thZ2UiLCJkaXJlY3Rpb24iLCJ1c2VTdHlsZU9wdGlvbnMiLCJzdHlsZU9wdGlvbnMiLCJsb2NhbGl6ZSIsInRheFRleHQiLCJ0b3RhbFRleHQiLCJ2YXRUZXh0IiwiYnVpbHRDYXJkIiwidXNlTWVtbyIsImJ1aWxkZXIiLCJBZGFwdGl2ZUNhcmRCdWlsZGVyIiwiSG9yaXpvbnRhbEFsaWdubWVudCIsIlRleHRTaXplIiwiVGV4dFdlaWdodCIsImJ1dHRvbnMiLCJmYWN0cyIsIml0ZW1zIiwidGF4IiwidGl0bGUiLCJ0b3RhbCIsInZhdCIsInJpY2hDYXJkV3JhcFRpdGxlIiwiYWRkVGV4dEJsb2NrIiwic2l6ZSIsIk1lZGl1bSIsIndlaWdodCIsIkJvbGRlciIsIndyYXAiLCJhZGRDb2x1bW5TZXQiLCJmaXJzdEZhY3RDb2x1bW4iLCJsYXN0RmFjdENvbHVtbiIsIm1hcCIsImtleSIsInZhbHVlIiwiaG9yaXpvbnRhbEFsaWdubWVudCIsIlJpZ2h0IiwiaW1hZ2UiLCJhbHQiLCJpbWFnZVRhcCIsInRhcCIsInVybCIsInByaWNlIiwicXVhbnRpdHkiLCJzdWJ0aXRsZSIsInRleHQiLCJpdGVtQ29sdW1ucyIsIml0ZW1JbWFnZUNvbHVtbiIsImNvbHVtbnMiLCJhZGRJbWFnZSIsInVuZGVmaW5lZCIsIml0ZW1UaXRsZUNvbHVtbiIsIml0ZW1QcmljZUNvbHVtbiIsInZhdENvbCIsInRheENvbCIsInRvdGFsQ29sIiwiYWRkQnV0dG9ucyIsImNhcmQiLCJkZWZhdWx0UHJvcHMiLCJwcm9wVHlwZXMiLCJQcm9wVHlwZXMiLCJzdHJpbmciLCJzaGFwZSIsImFycmF5IiwiYXJyYXlPZiIsImlzUmVxdWlyZWQiLCJhbnkiLCJib29sIl0sInNvdXJjZVJvb3QiOiJidW5kbGU6Ly8vIiwic291cmNlcyI6WyIuLi8uLi8uLi9zcmMvYWRhcHRpdmVDYXJkcy9BdHRhY2htZW50L1JlY2VpcHRDYXJkQ29udGVudC50c3giXSwic291cmNlc0NvbnRlbnQiOlsiLyogZXNsaW50IG5vLW1hZ2ljLW51bWJlcnM6IFtcImVycm9yXCIsIHsgXCJpZ25vcmVcIjogWzAsIDEsIDEwLCAxNSwgMjUsIDUwLCA3NV0gfV0gKi9cblxuaW1wb3J0IHsgaG9va3MgfSBmcm9tICdib3RmcmFtZXdvcmstd2ViY2hhdC1jb21wb25lbnQnO1xuaW1wb3J0IFByb3BUeXBlcyBmcm9tICdwcm9wLXR5cGVzJztcbmltcG9ydCBSZWFjdCwgeyBGQywgdXNlTWVtbyB9IGZyb20gJ3JlYWN0JztcbmltcG9ydCB0eXBlIHsgRGlyZWN0TGluZVJlY2VpcHRDYXJkIH0gZnJvbSAnYm90ZnJhbWV3b3JrLXdlYmNoYXQtY29yZSc7XG5cbmltcG9ydCBBZGFwdGl2ZUNhcmRCdWlsZGVyIGZyb20gJy4vQWRhcHRpdmVDYXJkQnVpbGRlcic7XG5pbXBvcnQgQWRhcHRpdmVDYXJkUmVuZGVyZXIgZnJvbSAnLi9BZGFwdGl2ZUNhcmRSZW5kZXJlcic7XG5pbXBvcnQgdXNlQWRhcHRpdmVDYXJkc1BhY2thZ2UgZnJvbSAnLi4vaG9va3MvdXNlQWRhcHRpdmVDYXJkc1BhY2thZ2UnO1xuaW1wb3J0IHVzZVN0eWxlT3B0aW9ucyBmcm9tICcuLi8uLi9ob29rcy91c2VTdHlsZU9wdGlvbnMnO1xuXG5jb25zdCB7IHVzZURpcmVjdGlvbiwgdXNlTG9jYWxpemVyIH0gPSBob29rcztcblxuZnVuY3Rpb24gbnVsbE9yVW5kZWZpbmVkKG9iaikge1xuICByZXR1cm4gb2JqID09PSBudWxsIHx8IHR5cGVvZiBvYmogPT09ICd1bmRlZmluZWQnO1xufVxuXG50eXBlIFJlY2VpcHRDYXJkQ29udGVudFByb3BzID0ge1xuICBhY3Rpb25QZXJmb3JtZWRDbGFzc05hbWU/OiBzdHJpbmc7XG4gIGNvbnRlbnQ6IERpcmVjdExpbmVSZWNlaXB0Q2FyZDtcbiAgZGlzYWJsZWQ/OiBib29sZWFuO1xufTtcblxuY29uc3QgUmVjZWlwdENhcmRDb250ZW50OiBGQzxSZWNlaXB0Q2FyZENvbnRlbnRQcm9wcz4gPSAoeyBhY3Rpb25QZXJmb3JtZWRDbGFzc05hbWUsIGNvbnRlbnQsIGRpc2FibGVkIH0pID0+IHtcbiAgY29uc3QgW2FkYXB0aXZlQ2FyZHNQYWNrYWdlXSA9IHVzZUFkYXB0aXZlQ2FyZHNQYWNrYWdlKCk7XG4gIGNvbnN0IFtkaXJlY3Rpb25dID0gdXNlRGlyZWN0aW9uKCk7XG4gIGNvbnN0IFtzdHlsZU9wdGlvbnNdID0gdXNlU3R5bGVPcHRpb25zKCk7XG4gIGNvbnN0IGxvY2FsaXplID0gdXNlTG9jYWxpemVyKCk7XG5cbiAgY29uc3QgdGF4VGV4dCA9IGxvY2FsaXplKCdSRUNFSVBUX0NBUkRfVEFYJyk7XG4gIGNvbnN0IHRvdGFsVGV4dCA9IGxvY2FsaXplKCdSRUNFSVBUX0NBUkRfVE9UQUwnKTtcbiAgY29uc3QgdmF0VGV4dCA9IGxvY2FsaXplKCdSRUNFSVBUX0NBUkRfVkFUJyk7XG5cbiAgY29uc3QgYnVpbHRDYXJkID0gdXNlTWVtbygoKSA9PiB7XG4gICAgY29uc3QgYnVpbGRlciA9IG5ldyBBZGFwdGl2ZUNhcmRCdWlsZGVyKGFkYXB0aXZlQ2FyZHNQYWNrYWdlLCBzdHlsZU9wdGlvbnMsIGRpcmVjdGlvbik7XG4gICAgY29uc3QgeyBIb3Jpem9udGFsQWxpZ25tZW50LCBUZXh0U2l6ZSwgVGV4dFdlaWdodCB9ID0gYWRhcHRpdmVDYXJkc1BhY2thZ2U7XG4gICAgY29uc3QgeyBidXR0b25zLCBmYWN0cywgaXRlbXMsIHRheCwgdGl0bGUsIHRvdGFsLCB2YXQgfSA9IGNvbnRlbnQ7XG4gICAgY29uc3QgeyByaWNoQ2FyZFdyYXBUaXRsZSB9ID0gc3R5bGVPcHRpb25zO1xuXG4gICAgaWYgKGNvbnRlbnQpIHtcbiAgICAgIGJ1aWxkZXIuYWRkVGV4dEJsb2NrKHRpdGxlLCB7IHNpemU6IFRleHRTaXplLk1lZGl1bSwgd2VpZ2h0OiBUZXh0V2VpZ2h0LkJvbGRlciwgd3JhcDogcmljaENhcmRXcmFwVGl0bGUgfSk7XG5cbiAgICAgIGlmIChmYWN0cykge1xuICAgICAgICBjb25zdCBbZmlyc3RGYWN0Q29sdW1uLCBsYXN0RmFjdENvbHVtbl0gPSBidWlsZGVyLmFkZENvbHVtblNldChbNzUsIDI1XSk7XG5cbiAgICAgICAgZmFjdHMubWFwKCh7IGtleSwgdmFsdWUgfSkgPT4ge1xuICAgICAgICAgIGJ1aWxkZXIuYWRkVGV4dEJsb2NrKGtleSwgeyBzaXplOiBUZXh0U2l6ZS5NZWRpdW0gfSwgZmlyc3RGYWN0Q29sdW1uKTtcbiAgICAgICAgICBidWlsZGVyLmFkZFRleHRCbG9jayhcbiAgICAgICAgICAgIHZhbHVlLFxuICAgICAgICAgICAgeyBzaXplOiBUZXh0U2l6ZS5NZWRpdW0sIGhvcml6b250YWxBbGlnbm1lbnQ6IEhvcml6b250YWxBbGlnbm1lbnQuUmlnaHQgfSxcbiAgICAgICAgICAgIGxhc3RGYWN0Q29sdW1uXG4gICAgICAgICAgKTtcbiAgICAgICAgfSk7XG4gICAgICB9XG5cbiAgICAgIGl0ZW1zICYmXG4gICAgICAgIGl0ZW1zLm1hcCgoeyBpbWFnZTogeyBhbHQsIHRhcDogaW1hZ2VUYXAsIHVybCB9ID0ge30sIHByaWNlLCBxdWFudGl0eSwgc3VidGl0bGUsIHRhcCwgdGV4dCwgdGl0bGUgfSkgPT4ge1xuICAgICAgICAgIGxldCBpdGVtQ29sdW1ucztcblxuICAgICAgICAgIGlmICh1cmwpIHtcbiAgICAgICAgICAgIGNvbnN0IFtpdGVtSW1hZ2VDb2x1bW4sIC4uLmNvbHVtbnNdID0gYnVpbGRlci5hZGRDb2x1bW5TZXQoWzE1LCA3NSwgMTBdKTtcblxuICAgICAgICAgICAgaXRlbUNvbHVtbnMgPSBjb2x1bW5zO1xuICAgICAgICAgICAgYnVpbGRlci5hZGRJbWFnZSh1cmwsIGl0ZW1JbWFnZUNvbHVtbiwgaW1hZ2VUYXAsIGFsdCk7XG4gICAgICAgICAgfSBlbHNlIHtcbiAgICAgICAgICAgIGl0ZW1Db2x1bW5zID0gYnVpbGRlci5hZGRDb2x1bW5TZXQoWzc1LCAyNV0sIHVuZGVmaW5lZCwgdGFwICYmIHRhcCk7XG4gICAgICAgICAgfVxuXG4gICAgICAgICAgY29uc3QgW2l0ZW1UaXRsZUNvbHVtbiwgaXRlbVByaWNlQ29sdW1uXSA9IGl0ZW1Db2x1bW5zO1xuXG4gICAgICAgICAgYnVpbGRlci5hZGRUZXh0QmxvY2soXG4gICAgICAgICAgICBxdWFudGl0eSA/IGAke3RpdGxlfSAmdGltZXM7ICR7cXVhbnRpdHl9YCA6IHRpdGxlLFxuICAgICAgICAgICAgeyBzaXplOiBUZXh0U2l6ZS5NZWRpdW0sIHdlaWdodDogVGV4dFdlaWdodC5Cb2xkZXIsIHdyYXA6IHJpY2hDYXJkV3JhcFRpdGxlIH0sXG4gICAgICAgICAgICBpdGVtVGl0bGVDb2x1bW5cbiAgICAgICAgICApO1xuICAgICAgICAgIGJ1aWxkZXIuYWRkVGV4dEJsb2NrKHN1YnRpdGxlLCB7IHNpemU6IFRleHRTaXplLk1lZGl1bSwgd3JhcDogcmljaENhcmRXcmFwVGl0bGUgfSwgaXRlbVRpdGxlQ29sdW1uKTtcbiAgICAgICAgICBidWlsZGVyLmFkZFRleHRCbG9jayhwcmljZSwgeyBob3Jpem9udGFsQWxpZ25tZW50OiBIb3Jpem9udGFsQWxpZ25tZW50LlJpZ2h0IH0sIGl0ZW1QcmljZUNvbHVtbik7XG5cbiAgICAgICAgICBpZiAodGV4dCkge1xuICAgICAgICAgICAgYnVpbGRlci5hZGRUZXh0QmxvY2sodGV4dCwgeyBzaXplOiBUZXh0U2l6ZS5NZWRpdW0sIHdyYXA6IHJpY2hDYXJkV3JhcFRpdGxlIH0sIGl0ZW1UaXRsZUNvbHVtbik7XG4gICAgICAgICAgfVxuICAgICAgICB9KTtcblxuICAgICAgaWYgKCFudWxsT3JVbmRlZmluZWQodmF0KSkge1xuICAgICAgICBjb25zdCB2YXRDb2wgPSBidWlsZGVyLmFkZENvbHVtblNldChbNzUsIDI1XSk7XG5cbiAgICAgICAgYnVpbGRlci5hZGRUZXh0QmxvY2sodmF0VGV4dCwgeyBzaXplOiBUZXh0U2l6ZS5NZWRpdW0sIHdlaWdodDogVGV4dFdlaWdodC5Cb2xkZXIgfSwgdmF0Q29sWzBdKTtcbiAgICAgICAgYnVpbGRlci5hZGRUZXh0QmxvY2sodmF0LCB7IGhvcml6b250YWxBbGlnbm1lbnQ6IEhvcml6b250YWxBbGlnbm1lbnQuUmlnaHQgfSwgdmF0Q29sWzFdKTtcbiAgICAgIH1cblxuICAgICAgaWYgKCFudWxsT3JVbmRlZmluZWQodGF4KSkge1xuICAgICAgICBjb25zdCB0YXhDb2wgPSBidWlsZGVyLmFkZENvbHVtblNldChbNzUsIDI1XSk7XG5cbiAgICAgICAgYnVpbGRlci5hZGRUZXh0QmxvY2sodGF4VGV4dCwgeyBzaXplOiBUZXh0U2l6ZS5NZWRpdW0sIHdlaWdodDogVGV4dFdlaWdodC5Cb2xkZXIgfSwgdGF4Q29sWzBdKTtcbiAgICAgICAgYnVpbGRlci5hZGRUZXh0QmxvY2sodGF4LCB7IGhvcml6b250YWxBbGlnbm1lbnQ6IEhvcml6b250YWxBbGlnbm1lbnQuUmlnaHQgfSwgdGF4Q29sWzFdKTtcbiAgICAgIH1cblxuICAgICAgaWYgKCFudWxsT3JVbmRlZmluZWQodG90YWwpKSB7XG4gICAgICAgIGNvbnN0IHRvdGFsQ29sID0gYnVpbGRlci5hZGRDb2x1bW5TZXQoWzc1LCAyNV0pO1xuXG4gICAgICAgIGJ1aWxkZXIuYWRkVGV4dEJsb2NrKHRvdGFsVGV4dCwgeyBzaXplOiBUZXh0U2l6ZS5NZWRpdW0sIHdlaWdodDogVGV4dFdlaWdodC5Cb2xkZXIgfSwgdG90YWxDb2xbMF0pO1xuICAgICAgICBidWlsZGVyLmFkZFRleHRCbG9jayhcbiAgICAgICAgICB0b3RhbCxcbiAgICAgICAgICB7IGhvcml6b250YWxBbGlnbm1lbnQ6IEhvcml6b250YWxBbGlnbm1lbnQuUmlnaHQsIHNpemU6IFRleHRTaXplLk1lZGl1bSwgd2VpZ2h0OiBUZXh0V2VpZ2h0LkJvbGRlciB9LFxuICAgICAgICAgIHRvdGFsQ29sWzFdXG4gICAgICAgICk7XG4gICAgICB9XG5cbiAgICAgIGJ1aWxkZXIuYWRkQnV0dG9ucyhidXR0b25zKTtcblxuICAgICAgcmV0dXJuIGJ1aWxkZXIuY2FyZDtcbiAgICB9XG4gIH0sIFthZGFwdGl2ZUNhcmRzUGFja2FnZSwgY29udGVudCwgZGlyZWN0aW9uLCBzdHlsZU9wdGlvbnMsIHRheFRleHQsIHRvdGFsVGV4dCwgdmF0VGV4dF0pO1xuXG4gIHJldHVybiAoXG4gICAgPEFkYXB0aXZlQ2FyZFJlbmRlcmVyXG4gICAgICBhY3Rpb25QZXJmb3JtZWRDbGFzc05hbWU9e2FjdGlvblBlcmZvcm1lZENsYXNzTmFtZX1cbiAgICAgIGFkYXB0aXZlQ2FyZD17YnVpbHRDYXJkfVxuICAgICAgZGlzYWJsZWQ9e2Rpc2FibGVkfVxuICAgICAgdGFwQWN0aW9uPXtjb250ZW50ICYmIGNvbnRlbnQudGFwfVxuICAgIC8+XG4gICk7XG59O1xuXG5SZWNlaXB0Q2FyZENvbnRlbnQuZGVmYXVsdFByb3BzID0ge1xuICBhY3Rpb25QZXJmb3JtZWRDbGFzc05hbWU6ICcnLFxuICBkaXNhYmxlZDogdW5kZWZpbmVkXG59O1xuXG5SZWNlaXB0Q2FyZENvbnRlbnQucHJvcFR5cGVzID0ge1xuICBhY3Rpb25QZXJmb3JtZWRDbGFzc05hbWU6IFByb3BUeXBlcy5zdHJpbmcsXG4gIC8vIFByb3BUeXBlcyBjYW5ub3QgZnVsbHkgY2FwdHVyZSBUeXBlU2NyaXB0IHR5cGVzLlxuICAvLyBAdHMtaWdub3JlXG4gIGNvbnRlbnQ6IFByb3BUeXBlcy5zaGFwZSh7XG4gICAgYnV0dG9uczogUHJvcFR5cGVzLmFycmF5LFxuICAgIGZhY3RzOiBQcm9wVHlwZXMuYXJyYXlPZihcbiAgICAgIFByb3BUeXBlcy5zaGFwZSh7XG4gICAgICAgIGtleTogUHJvcFR5cGVzLnN0cmluZyxcbiAgICAgICAgdmFsdWU6IFByb3BUeXBlcy5zdHJpbmdcbiAgICAgIH0pXG4gICAgKSxcbiAgICBpdGVtczogUHJvcFR5cGVzLmFycmF5T2YoXG4gICAgICBQcm9wVHlwZXMuc2hhcGUoe1xuICAgICAgICBpbWFnZTogUHJvcFR5cGVzLnNoYXBlKHtcbiAgICAgICAgICBhbHQ6IFByb3BUeXBlcy5zdHJpbmcuaXNSZXF1aXJlZCxcbiAgICAgICAgICB0YXA6IFByb3BUeXBlcy5hbnksXG4gICAgICAgICAgdXJsOiBQcm9wVHlwZXMuc3RyaW5nLmlzUmVxdWlyZWRcbiAgICAgICAgfSksXG4gICAgICAgIHByaWNlOiBQcm9wVHlwZXMuc3RyaW5nLmlzUmVxdWlyZWQsXG4gICAgICAgIHF1YW50aXR5OiBQcm9wVHlwZXMuc3RyaW5nLFxuICAgICAgICBzdWJ0aXRsZTogUHJvcFR5cGVzLnN0cmluZyxcbiAgICAgICAgdGFwOiBQcm9wVHlwZXMuYW55LFxuICAgICAgICB0ZXh0OiBQcm9wVHlwZXMuc3RyaW5nLFxuICAgICAgICB0aXRsZTogUHJvcFR5cGVzLnN0cmluZy5pc1JlcXVpcmVkXG4gICAgICB9KVxuICAgICksXG4gICAgdGFwOiBQcm9wVHlwZXMuYW55LFxuICAgIHRheDogUHJvcFR5cGVzLnN0cmluZyxcbiAgICB0aXRsZTogUHJvcFR5cGVzLnN0cmluZyxcbiAgICB0b3RhbDogUHJvcFR5cGVzLnN0cmluZyxcbiAgICB2YXQ6IFByb3BUeXBlcy5zdHJpbmdcbiAgfSkuaXNSZXF1aXJlZCxcbiAgZGlzYWJsZWQ6IFByb3BUeXBlcy5ib29sXG59O1xuXG5leHBvcnQgZGVmYXVsdCBSZWNlaXB0Q2FyZENvbnRlbnQ7XG4iXSwibWFwcGluZ3MiOiI7Ozs7Ozs7Ozs7Ozs7OztBQUVBOztBQUNBOztBQUNBOztBQUdBOztBQUNBOztBQUNBOztBQUNBOzs7Ozs7QUFWQTtBQVlBLElBQVFBLFlBQVIsR0FBdUNDLG1DQUF2QyxDQUFRRCxZQUFSO0FBQUEsSUFBc0JFLFlBQXRCLEdBQXVDRCxtQ0FBdkMsQ0FBc0JDLFlBQXRCOztBQUVBLFNBQVNDLGVBQVQsQ0FBeUJDLEdBQXpCLEVBQThC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