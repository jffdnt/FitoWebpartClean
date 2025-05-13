"use strict";

var _interopRequireDefault = require("@babel/runtime/helpers/interopRequireDefault");

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.default = void 0;

var _propTypes = _interopRequireDefault(require("prop-types"));

var _react = _interopRequireDefault(require("react"));

var _botframeworkWebchatComponent = require("botframework-webchat-component");

/* eslint-disable react/no-array-index-key */
var useLocalizer = _botframeworkWebchatComponent.hooks.useLocalizer;

var RichCardAttachment = function RichCardAttachment(_ref) {
  var _ref$content = _ref.content,
      content = _ref$content === void 0 ? {} : _ref$content;
  var localize = useLocalizer();
  var buttons = content.buttons,
      facts = content.facts,
      image = content.image,
      images = content.images,
      items = content.items,
      subtitle = content.subtitle,
      tax = content.tax,
      text = content.text,
      title = content.title,
      total = content.total,
      vat = content.vat;
  var taxLabel = localize('RECEIPT_CARD_TAX');
  var totalLabel = localize('RECEIPT_CARD_TOTAL');
  var vatLabel = localize('RECEIPT_CARD_VAT');
  var cardLabel = localize('ATTACHMENT_CARD', title || '', subtitle || '', text || '');
  return /*#__PURE__*/_react.default.createElement("article", null, /*#__PURE__*/_react.default.createElement("div", null, cardLabel), !!image && !!image.alt && /*#__PURE__*/_react.default.createElement("img", {
    alt: image.alt
  }), !!images && !!images.length && images.map(function (_ref2, index) {
    var alt = _ref2.alt;
    return /*#__PURE__*/_react.default.createElement("img", {
      alt: alt,
      key: index
    });
  }), !!facts && !!facts.length && /*#__PURE__*/_react.default.createElement("dl", null, facts.map(function (_ref3, index) {
    var key = _ref3.key,
        value = _ref3.value;
    return /*#__PURE__*/_react.default.createElement(_react.default.Fragment, {
      key: index
    }, /*#__PURE__*/_react.default.createElement("dt", null, key), /*#__PURE__*/_react.default.createElement("dd", null, value));
  })), !!items && !!items.length && /*#__PURE__*/_react.default.createElement("ul", null, items.map(function (_ref4, index) {
    var image = _ref4.image,
        price = _ref4.price,
        quantity = _ref4.quantity,
        subtitle = _ref4.subtitle,
        text = _ref4.text,
        title = _ref4.title;
    return /*#__PURE__*/_react.default.createElement("li", {
      key: index
    }, !!image && !!image.alt && /*#__PURE__*/_react.default.createElement("img", {
      alt: image.alt
    }), !!title && /*#__PURE__*/_react.default.createElement("p", null, title), !!subtitle && /*#__PURE__*/_react.default.createElement("p", null, subtitle), !!text && /*#__PURE__*/_react.default.createElement("p", null, text), !!quantity && /*#__PURE__*/_react.default.createElement("p", null, quantity), !!price && /*#__PURE__*/_react.default.createElement("p", null, price));
  })), !!vat && /*#__PURE__*/_react.default.createElement("p", null, vatLabel, " ", vat), !!tax && /*#__PURE__*/_react.default.createElement("p", null, taxLabel, " ", tax), !!total && /*#__PURE__*/_react.default.createElement("p", null, totalLabel, " ", total), !!buttons && !!buttons.length && /*#__PURE__*/_react.default.createElement("div", null, buttons.map(function (_ref5, index) {
    var title = _ref5.title;
    return /*#__PURE__*/_react.default.createElement("button", {
      key: index,
      tabIndex: -1,
      type: "button"
    }, title);
  })));
};

RichCardAttachment.propTypes = {
  content: _propTypes.default.any.isRequired
};
var _default = RichCardAttachment;
exports.default = _default;
//# sourceMappingURL=data:application/json;charset=utf-8;base64,eyJ2ZXJzaW9uIjozLCJuYW1lcyI6WyJ1c2VMb2NhbGl6ZXIiLCJob29rcyIsIlJpY2hDYXJkQXR0YWNobWVudCIsImNvbnRlbnQiLCJsb2NhbGl6ZSIsImJ1dHRvbnMiLCJmYWN0cyIsImltYWdlIiwiaW1hZ2VzIiwiaXRlbXMiLCJzdWJ0aXRsZSIsInRheCIsInRleHQiLCJ0aXRsZSIsInRvdGFsIiwidmF0IiwidGF4TGFiZWwiLCJ0b3RhbExhYmVsIiwidmF0TGFiZWwiLCJjYXJkTGFiZWwiLCJhbHQiLCJsZW5ndGgiLCJtYXAiLCJpbmRleCIsImtleSIsInZhbHVlIiwicHJpY2UiLCJxdWFudGl0eSIsInByb3BUeXBlcyIsIlByb3BUeXBlcyIsImFueSIsImlzUmVxdWlyZWQiXSwic291cmNlUm9vdCI6ImJ1bmRsZTovLy8iLCJzb3VyY2VzIjpbIi4uLy4uLy4uL3NyYy9hZGFwdGl2ZUNhcmRzL0F0dGFjaG1lbnRGb3JTY3JlZW5SZWFkZXIvUmljaENhcmRBdHRhY2htZW50LmpzIl0sInNvdXJjZXNDb250ZW50IjpbIi8qIGVzbGludC1kaXNhYmxlIHJlYWN0L25vLWFycmF5LWluZGV4LWtleSAqL1xuaW1wb3J0IFByb3BUeXBlcyBmcm9tICdwcm9wLXR5cGVzJztcbmltcG9ydCBSZWFjdCBmcm9tICdyZWFjdCc7XG5pbXBvcnQgeyBob29rcyB9IGZyb20gJ2JvdGZyYW1ld29yay13ZWJjaGF0LWNvbXBvbmVudCc7XG5cbmNvbnN0IHsgdXNlTG9jYWxpemVyIH0gPSBob29rcztcblxuY29uc3QgUmljaENhcmRBdHRhY2htZW50ID0gKHsgY29udGVudCA9IHt9IH0pID0+IHtcbiAgY29uc3QgbG9jYWxpemUgPSB1c2VMb2NhbGl6ZXIoKTtcbiAgY29uc3QgeyBidXR0b25zLCBmYWN0cywgaW1hZ2UsIGltYWdlcywgaXRlbXMsIHN1YnRpdGxlLCB0YXgsIHRleHQsIHRpdGxlLCB0b3RhbCwgdmF0IH0gPSBjb250ZW50O1xuXG4gIGNvbnN0IHRheExhYmVsID0gbG9jYWxpemUoJ1JFQ0VJUFRfQ0FSRF9UQVgnKTtcbiAgY29uc3QgdG90YWxMYWJlbCA9IGxvY2FsaXplKCdSRUNFSVBUX0NBUkRfVE9UQUwnKTtcbiAgY29uc3QgdmF0TGFiZWwgPSBsb2NhbGl6ZSgnUkVDRUlQVF9DQVJEX1ZBVCcpO1xuXG4gIGNvbnN0IGNhcmRMYWJlbCA9IGxvY2FsaXplKCdBVFRBQ0hNRU5UX0NBUkQnLCB0aXRsZSB8fCAnJywgc3VidGl0bGUgfHwgJycsIHRleHQgfHwgJycpO1xuXG4gIHJldHVybiAoXG4gICAgPGFydGljbGU+XG4gICAgICA8ZGl2PntjYXJkTGFiZWx9PC9kaXY+XG4gICAgICB7ISFpbWFnZSAmJiAhIWltYWdlLmFsdCAmJiA8aW1nIGFsdD17aW1hZ2UuYWx0fSAvPn1cbiAgICAgIHshIWltYWdlcyAmJiAhIWltYWdlcy5sZW5ndGggJiYgaW1hZ2VzLm1hcCgoeyBhbHQgfSwgaW5kZXgpID0+IDxpbWcgYWx0PXthbHR9IGtleT17aW5kZXh9IC8+KX1cbiAgICAgIHshIWZhY3RzICYmICEhZmFjdHMubGVuZ3RoICYmIChcbiAgICAgICAgPGRsPlxuICAgICAgICAgIHtmYWN0cy5tYXAoKHsga2V5LCB2YWx1ZSB9LCBpbmRleCkgPT4gKFxuICAgICAgICAgICAgPFJlYWN0LkZyYWdtZW50IGtleT17aW5kZXh9PlxuICAgICAgICAgICAgICA8ZHQ+e2tleX08L2R0PlxuICAgICAgICAgICAgICA8ZGQ+e3ZhbHVlfTwvZGQ+XG4gICAgICAgICAgICA8L1JlYWN0LkZyYWdtZW50PlxuICAgICAgICAgICkpfVxuICAgICAgICA8L2RsPlxuICAgICAgKX1cbiAgICAgIHshIWl0ZW1zICYmICEhaXRlbXMubGVuZ3RoICYmIChcbiAgICAgICAgPHVsPlxuICAgICAgICAgIHtpdGVtcy5tYXAoKHsgaW1hZ2UsIHByaWNlLCBxdWFudGl0eSwgc3VidGl0bGUsIHRleHQsIHRpdGxlIH0sIGluZGV4KSA9PiAoXG4gICAgICAgICAgICA8bGkga2V5PXtpbmRleH0+XG4gICAgICAgICAgICAgIHshIWltYWdlICYmICEhaW1hZ2UuYWx0ICYmIDxpbWcgYWx0PXtpbWFnZS5hbHR9IC8+fVxuICAgICAgICAgICAgICB7ISF0aXRsZSAmJiA8cD57dGl0bGV9PC9wPn1cbiAgICAgICAgICAgICAgeyEhc3VidGl0bGUgJiYgPHA+e3N1YnRpdGxlfTwvcD59XG4gICAgICAgICAgICAgIHshIXRleHQgJiYgPHA+e3RleHR9PC9wPn1cbiAgICAgICAgICAgICAgeyEhcXVhbnRpdHkgJiYgPHA+e3F1YW50aXR5fTwvcD59XG4gICAgICAgICAgICAgIHshIXByaWNlICYmIDxwPntwcmljZX08L3A+fVxuICAgICAgICAgICAgPC9saT5cbiAgICAgICAgICApKX1cbiAgICAgICAgPC91bD5cbiAgICAgICl9XG4gICAgICB7ISF2YXQgJiYgKFxuICAgICAgICA8cD5cbiAgICAgICAgICB7dmF0TGFiZWx9IHt2YXR9XG4gICAgICAgIDwvcD5cbiAgICAgICl9XG4gICAgICB7ISF0YXggJiYgKFxuICAgICAgICA8cD5cbiAgICAgICAgICB7dGF4TGFiZWx9IHt0YXh9XG4gICAgICAgIDwvcD5cbiAgICAgICl9XG4gICAgICB7ISF0b3RhbCAmJiAoXG4gICAgICAgIDxwPlxuICAgICAgICAgIHt0b3RhbExhYmVsfSB7dG90YWx9XG4gICAgICAgIDwvcD5cbiAgICAgICl9XG4gICAgICB7ISFidXR0b25zICYmICEhYnV0dG9ucy5sZW5ndGggJiYgKFxuICAgICAgICA8ZGl2PlxuICAgICAgICAgIHtidXR0b25zLm1hcCgoeyB0aXRsZSB9LCBpbmRleCkgPT4gKFxuICAgICAgICAgICAgPGJ1dHRvbiBrZXk9e2luZGV4fSB0YWJJbmRleD17LTF9IHR5cGU9XCJidXR0b25cIj5cbiAgICAgICAgICAgICAge3RpdGxlfVxuICAgICAgICAgICAgPC9idXR0b24+XG4gICAgICAgICAgKSl9XG4gICAgICAgIDwvZGl2PlxuICAgICAgKX1cbiAgICA8L2FydGljbGU+XG4gICk7XG59O1xuXG5SaWNoQ2FyZEF0dGFjaG1lbnQucHJvcFR5cGVzID0ge1xuICBjb250ZW50OiBQcm9wVHlwZXMuYW55LmlzUmVxdWlyZWRcbn07XG5cbmV4cG9ydCBkZWZhdWx0IFJpY2hDYXJkQXR0YWNobWVudDtcbiJdLCJtYXBwaW5ncyI6Ijs7Ozs7Ozs7O0FBQ0E7O0FBQ0E7O0FBQ0E7O0FBSEE7QUFLQSxJQUFRQSxZQUFSLEdBQXlCQyxtQ0FBekIsQ0FBUUQsWUFBUjs7QUFFQSxJQUFNRSxrQkFBa0IsR0FBRyxTQUFyQkEsa0JBQXFCLE9BQXNCO0VBQUEsd0JBQW5CQyxPQUFtQjtFQUFBLElBQW5CQSxPQUFtQiw2QkFBVCxFQUFTO0VBQy9DLElBQU1DLFFBQVEsR0FBR0osWUFBWSxFQUE3QjtFQUNBLElBQVFLLE9BQVIsR0FBeUZGLE9BQXpGLENBQVFFLE9BQVI7RUFBQSxJQUFpQkMsS0FBakIsR0FBeUZILE9BQXpGLENBQWlCRyxLQUFqQjtFQUFBLElBQXdCQyxLQUF4QixHQUF5RkosT0FBekYsQ0FBd0JJLEtBQXhCO0VBQUEsSUFBK0JDLE1BQS9CLEdBQXlGTCxPQUF6RixDQUErQkssTUFBL0I7RUFBQSxJQUF1Q0MsS0FBdkMsR0FBeUZOLE9BQXpGLENBQXVDTSxLQUF2QztFQUFBLElBQThDQyxRQUE5QyxHQUF5RlAsT0FBekYsQ0FBOENPLFFBQTlDO0VBQUEsSUFBd0RDLEdBQXhELEdBQXlGUixPQUF6RixDQUF3RFEsR0FBeEQ7RUFBQSxJQUE2REMsSUFBN0QsR0FBeUZULE9BQXpGLENBQTZEUyxJQUE3RDtFQUFBLElBQW1FQyxLQUFuRSxHQUF5RlYsT0FBekYsQ0FBbUVVLEtBQW5FO0VBQUEsSUFBMEVDLEtBQTFFLEdBQXlGWCxPQUF6RixDQUEwRVcsS0FBMUU7RUFBQSxJQUFpRkMsR0FBakYsR0FBeUZaLE9BQXpGLENBQWlGWSxHQUFqRjtFQUVBLElBQU1DLFFBQVEsR0FBR1osUUFBUSxDQUFDLGtCQUFELENBQXpCO0VBQ0EsSUFBTWEsVUFBVSxHQUFHYixRQUFRLENBQUMsb0JBQUQsQ0FBM0I7RUFDQSxJQUFNYyxRQUFRLEdBQUdkLFFBQVEsQ0FBQyxrQkFB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