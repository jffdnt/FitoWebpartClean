"use strict";

var _interopRequireDefault = require("@babel/runtime/helpers/interopRequireDefault");

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.default = createAdaptiveCardsAttachmentMiddleware;

var _react = _interopRequireDefault(require("react"));

var _AdaptiveCardAttachment = _interopRequireDefault(require("./AttachmentForScreenReader/AdaptiveCardAttachment"));

var _RichCardAttachment = _interopRequireDefault(require("./AttachmentForScreenReader/RichCardAttachment"));

var RICH_CARD_CONTENT_TYPES = ['application/vnd.microsoft.card.animation', 'application/vnd.microsoft.card.audio', 'application/vnd.microsoft.card.hero', 'application/vnd.microsoft.card.oauth', 'application/vnd.microsoft.card.receipt', 'application/vnd.microsoft.card.signin', 'application/vnd.microsoft.card.thumbnail', 'application/vnd.microsoft.card.video'];

function createAdaptiveCardsAttachmentMiddleware() {
  return function () {
    return function (next) {
      return function () {
        for (var _len = arguments.length, args = new Array(_len), _key = 0; _key < _len; _key++) {
          args[_key] = arguments[_key];
        }

        var _args$0$attachment = args[0].attachment,
            content = _args$0$attachment.content,
            contentType = _args$0$attachment.contentType;
        return content && RICH_CARD_CONTENT_TYPES.includes(contentType) ? function () {
          return /*#__PURE__*/_react.default.createElement(_RichCardAttachment.default, {
            content: content
          });
        } : content && contentType === 'application/vnd.microsoft.card.adaptive' ? function () {
          return /*#__PURE__*/_react.default.createElement(_AdaptiveCardAttachment.default, {
            content: content
          });
        } : next.apply(void 0, args);
      };
    };
  };
}
//# sourceMappingURL=data:application/json;charset=utf-8;base64,eyJ2ZXJzaW9uIjozLCJuYW1lcyI6WyJSSUNIX0NBUkRfQ09OVEVOVF9UWVBFUyIsImNyZWF0ZUFkYXB0aXZlQ2FyZHNBdHRhY2htZW50TWlkZGxld2FyZSIsIm5leHQiLCJhcmdzIiwiYXR0YWNobWVudCIsImNvbnRlbnQiLCJjb250ZW50VHlwZSIsImluY2x1ZGVzIl0sInNvdXJjZVJvb3QiOiJidW5kbGU6Ly8vIiwic291cmNlcyI6WyIuLi8uLi9zcmMvYWRhcHRpdmVDYXJkcy9jcmVhdGVBZGFwdGl2ZUNhcmRzQXR0YWNobWVudEZvclNjcmVlblJlYWRlck1pZGRsZXdhcmUudHN4Il0sInNvdXJjZXNDb250ZW50IjpbImltcG9ydCB7IEF0dGFjaG1lbnRGb3JTY3JlZW5SZWFkZXJNaWRkbGV3YXJlIH0gZnJvbSAnYm90ZnJhbWV3b3JrLXdlYmNoYXQtYXBpJztcbmltcG9ydCBSZWFjdCBmcm9tICdyZWFjdCc7XG5cbmltcG9ydCBBZGFwdGl2ZUNhcmRBdHRhY2htZW50IGZyb20gJy4vQXR0YWNobWVudEZvclNjcmVlblJlYWRlci9BZGFwdGl2ZUNhcmRBdHRhY2htZW50JztcbmltcG9ydCBSaWNoQ2FyZEF0dGFjaG1lbnQgZnJvbSAnLi9BdHRhY2htZW50Rm9yU2NyZWVuUmVhZGVyL1JpY2hDYXJkQXR0YWNobWVudCc7XG5cbmNvbnN0IFJJQ0hfQ0FSRF9DT05URU5UX1RZUEVTID0gW1xuICAnYXBwbGljYXRpb24vdm5kLm1pY3Jvc29mdC5jYXJkLmFuaW1hdGlvbicsXG4gICdhcHBsaWNhdGlvbi92bmQubWljcm9zb2Z0LmNhcmQuYXVkaW8nLFxuICAnYXBwbGljYXRpb24vdm5kLm1pY3Jvc29mdC5jYXJkLmhlcm8nLFxuICAnYXBwbGljYXRpb24vdm5kLm1pY3Jvc29mdC5jYXJkLm9hdXRoJyxcbiAgJ2FwcGxpY2F0aW9uL3ZuZC5taWNyb3NvZnQuY2FyZC5yZWNlaXB0JyxcbiAgJ2FwcGxpY2F0aW9uL3ZuZC5taWNyb3NvZnQuY2FyZC5zaWduaW4nLFxuICAnYXBwbGljYXRpb24vdm5kLm1pY3Jvc29mdC5jYXJkLnRodW1ibmFpbCcsXG4gICdhcHBsaWNhdGlvbi92bmQubWljcm9zb2Z0LmNhcmQudmlkZW8nXG5dO1xuXG5leHBvcnQgZGVmYXVsdCBmdW5jdGlvbiBjcmVhdGVBZGFwdGl2ZUNhcmRzQXR0YWNobWVudE1pZGRsZXdhcmUoKTogQXR0YWNobWVudEZvclNjcmVlblJlYWRlck1pZGRsZXdhcmUge1xuICByZXR1cm4gKCkgPT5cbiAgICBuZXh0ID0+XG4gICAgKC4uLmFyZ3MpID0+IHtcbiAgICAgIGNvbnN0IFtcbiAgICAgICAge1xuICAgICAgICAgIGF0dGFjaG1lbnQ6IHsgY29udGVudCwgY29udGVudFR5cGUgfVxuICAgICAgICB9XG4gICAgICBdID0gYXJncztcblxuICAgICAgcmV0dXJuIGNvbnRlbnQgJiYgUklDSF9DQVJEX0NPTlRFTlRfVFlQRVMuaW5jbHVkZXMoY29udGVudFR5cGUpXG4gICAgICAgID8gKCkgPT4gPFJpY2hDYXJkQXR0YWNobWVudCBjb250ZW50PXtjb250ZW50fSAvPlxuICAgICAgICA6IGNvbnRlbnQgJiYgY29udGVudFR5cGUgPT09ICdhcHBsaWNhdGlvbi92bmQubWljcm9zb2Z0LmNhcmQuYWRhcHRpdmUnXG4gICAgICAgID8gKCkgPT4gPEFkYXB0aXZlQ2FyZEF0dGFjaG1lbnQgY29udGVudD17Y29udGVudH0gLz5cbiAgICAgICAgOiBuZXh0KC4uLmFyZ3MpO1xuICAgIH07XG59XG4iXSwibWFwcGluZ3MiOiI7Ozs7Ozs7OztBQUNBOztBQUVBOztBQUNBOztBQUVBLElBQU1BLHVCQUF1QixHQUFHLENBQzlCLDBDQUQ4QixFQUU5QixzQ0FGOEIsRUFHOUIscUNBSDhCLEVBSTlCLHNDQUo4QixFQUs5Qix3Q0FMOEIsRUFNOUIsdUNBTjhCLEVBTzlCLDBDQVA4QixFQVE5QixzQ0FSOEIsQ0FBaEM7O0FBV2UsU0FBU0MsdUNBQVQsR0FBd0Y7RUFDckcsT0FBTztJQUFBLE9BQ0wsVUFBQUMsSUFBSTtNQUFBLE9BQ0osWUFBYTtRQUFBLGtDQUFUQyxJQUFTO1VBQVRBLElBQVM7UUFBQTs7UUFDWCx5QkFJSUEsSUFKSixJQUVJQyxVQUZKO1FBQUEsSUFFa0JDLE9BRmxCLHNCQUVrQkEsT0FGbEI7UUFBQSxJQUUyQkMsV0FGM0Isc0JBRTJCQSxXQUYzQjtRQU1BLE9BQU9ELE9BQU8sSUFBSUwsdUJBQXVCLENBQUNPLFFBQXhCLENBQWlDRCxXQUFqQyxDQUFYLEdBQ0g7VUFBQSxvQkFBTSw2QkFBQywyQkFBRDtZQUFvQixPQUFPLEVBQUVEO1VBQTdCLEVBQU47UUFBQSxDQURHLEdBRUhBLE9BQU8sSUFBSUMsV0FBVyxLQUFLLHlDQUEzQixHQUNBO1VBQUEsb0JBQU0sNkJBQUMsK0JBQUQ7WUFBd0IsT0FBTyxFQUFFRDtVQUFqQyxFQUFOO1FBQUEsQ0FEQSxHQUVBSCxJQUFJLE1BQUosU0FBUUMsSUFBUixDQUpKO01BS0QsQ0FiRztJQUFBLENBREM7RUFBQSxDQUFQO0FBZUQifQ==