"use strict";

var _interopRequireDefault = require("@babel/runtime/helpers/interopRequireDefault");

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.default = createAdaptiveCardsAttachmentMiddleware;

var _react = _interopRequireDefault(require("react"));

var _AdaptiveCardAttachment = _interopRequireDefault(require("./Attachment/AdaptiveCardAttachment"));

var _AnimationCardAttachment = _interopRequireDefault(require("./Attachment/AnimationCardAttachment"));

var _AudioCardAttachment = _interopRequireDefault(require("./Attachment/AudioCardAttachment"));

var _HeroCardAttachment = _interopRequireDefault(require("./Attachment/HeroCardAttachment"));

var _OAuthCardAttachment = _interopRequireDefault(require("./Attachment/OAuthCardAttachment"));

var _ReceiptCardAttachment = _interopRequireDefault(require("./Attachment/ReceiptCardAttachment"));

var _SignInCardAttachment = _interopRequireDefault(require("./Attachment/SignInCardAttachment"));

var _ThumbnailCardAttachment = _interopRequireDefault(require("./Attachment/ThumbnailCardAttachment"));

var _VideoCardAttachment = _interopRequireDefault(require("./Attachment/VideoCardAttachment"));

function createAdaptiveCardsAttachmentMiddleware() {
  // This is not returning a React component, but a render function.
  return function () {
    return function (next) {
      return function () {
        for (var _len = arguments.length, args = new Array(_len), _key = 0; _key < _len; _key++) {
          args[_key] = arguments[_key];
        }

        var attachment = args[0].attachment;
        return attachment.contentType === 'application/vnd.microsoft.card.hero' ? /*#__PURE__*/_react.default.createElement(_HeroCardAttachment.default, {
          attachment: attachment
        }) : attachment.contentType === 'application/vnd.microsoft.card.adaptive' ? /*#__PURE__*/_react.default.createElement(_AdaptiveCardAttachment.default, {
          attachment: attachment
        }) : attachment.contentType === 'application/vnd.microsoft.card.animation' ? /*#__PURE__*/_react.default.createElement(_AnimationCardAttachment.default, {
          attachment: attachment
        }) : attachment.contentType === 'application/vnd.microsoft.card.audio' ? /*#__PURE__*/_react.default.createElement(_AudioCardAttachment.default, {
          attachment: attachment
        }) : attachment.contentType === 'application/vnd.microsoft.card.oauth' ? /*#__PURE__*/_react.default.createElement(_OAuthCardAttachment.default, {
          attachment: attachment
        }) : attachment.contentType === 'application/vnd.microsoft.card.receipt' ? /*#__PURE__*/_react.default.createElement(_ReceiptCardAttachment.default, {
          attachment: attachment
        }) : attachment.contentType === 'application/vnd.microsoft.card.signin' ? /*#__PURE__*/_react.default.createElement(_SignInCardAttachment.default, {
          attachment: attachment
        }) : attachment.contentType === 'application/vnd.microsoft.card.thumbnail' ? /*#__PURE__*/_react.default.createElement(_ThumbnailCardAttachment.default, {
          attachment: attachment
        }) : attachment.contentType === 'application/vnd.microsoft.card.video' ? /*#__PURE__*/_react.default.createElement(_VideoCardAttachment.default, {
          attachment: attachment
        }) : next.apply(void 0, args);
      };
    };
  };
}
//# sourceMappingURL=data:application/json;charset=utf-8;base64,eyJ2ZXJzaW9uIjozLCJuYW1lcyI6WyJjcmVhdGVBZGFwdGl2ZUNhcmRzQXR0YWNobWVudE1pZGRsZXdhcmUiLCJuZXh0IiwiYXJncyIsImF0dGFjaG1lbnQiLCJjb250ZW50VHlwZSJdLCJzb3VyY2VSb290IjoiYnVuZGxlOi8vLyIsInNvdXJjZXMiOlsiLi4vLi4vc3JjL2FkYXB0aXZlQ2FyZHMvY3JlYXRlQWRhcHRpdmVDYXJkc0F0dGFjaG1lbnRNaWRkbGV3YXJlLnRzeCJdLCJzb3VyY2VzQ29udGVudCI6WyJpbXBvcnQgUmVhY3QgZnJvbSAncmVhY3QnO1xuXG5pbXBvcnQgeyBBdHRhY2htZW50TWlkZGxld2FyZSB9IGZyb20gJ2JvdGZyYW1ld29yay13ZWJjaGF0LWFwaSc7XG5cbmltcG9ydCBBZGFwdGl2ZUNhcmRBdHRhY2htZW50IGZyb20gJy4vQXR0YWNobWVudC9BZGFwdGl2ZUNhcmRBdHRhY2htZW50JztcbmltcG9ydCBBbmltYXRpb25DYXJkQXR0YWNobWVudCBmcm9tICcuL0F0dGFjaG1lbnQvQW5pbWF0aW9uQ2FyZEF0dGFjaG1lbnQnO1xuaW1wb3J0IEF1ZGlvQ2FyZEF0dGFjaG1lbnQgZnJvbSAnLi9BdHRhY2htZW50L0F1ZGlvQ2FyZEF0dGFjaG1lbnQnO1xuaW1wb3J0IEhlcm9DYXJkQXR0YWNobWVudCBmcm9tICcuL0F0dGFjaG1lbnQvSGVyb0NhcmRBdHRhY2htZW50JztcbmltcG9ydCBPQXV0aENhcmRBdHRhY2htZW50IGZyb20gJy4vQXR0YWNobWVudC9PQXV0aENhcmRBdHRhY2htZW50JztcbmltcG9ydCBSZWNlaXB0Q2FyZEF0dGFjaG1lbnQgZnJvbSAnLi9BdHRhY2htZW50L1JlY2VpcHRDYXJkQXR0YWNobWVudCc7XG5pbXBvcnQgU2lnbkluQ2FyZEF0dGFjaG1lbnQgZnJvbSAnLi9BdHRhY2htZW50L1NpZ25JbkNhcmRBdHRhY2htZW50JztcbmltcG9ydCBUaHVtYm5haWxDYXJkQXR0YWNobWVudCBmcm9tICcuL0F0dGFjaG1lbnQvVGh1bWJuYWlsQ2FyZEF0dGFjaG1lbnQnO1xuaW1wb3J0IFZpZGVvQ2FyZEF0dGFjaG1lbnQgZnJvbSAnLi9BdHRhY2htZW50L1ZpZGVvQ2FyZEF0dGFjaG1lbnQnO1xuXG5leHBvcnQgZGVmYXVsdCBmdW5jdGlvbiBjcmVhdGVBZGFwdGl2ZUNhcmRzQXR0YWNobWVudE1pZGRsZXdhcmUoKTogQXR0YWNobWVudE1pZGRsZXdhcmUge1xuICAvLyBUaGlzIGlzIG5vdCByZXR1cm5pbmcgYSBSZWFjdCBjb21wb25lbnQsIGJ1dCBhIHJlbmRlciBmdW5jdGlvbi5cbiAgcmV0dXJuICgpID0+XG4gICAgbmV4dCA9PlxuICAgICguLi5hcmdzKSA9PiB7XG4gICAgICBjb25zdCBbeyBhdHRhY2htZW50IH1dID0gYXJncztcblxuICAgICAgcmV0dXJuIGF0dGFjaG1lbnQuY29udGVudFR5cGUgPT09ICdhcHBsaWNhdGlvbi92bmQubWljcm9zb2Z0LmNhcmQuaGVybycgPyAoXG4gICAgICAgIDxIZXJvQ2FyZEF0dGFjaG1lbnQgYXR0YWNobWVudD17YXR0YWNobWVudH0gLz5cbiAgICAgICkgOiBhdHRhY2htZW50LmNvbnRlbnRUeXBlID09PSAnYXBwbGljYXRpb24vdm5kLm1pY3Jvc29mdC5jYXJkLmFkYXB0aXZlJyA/IChcbiAgICAgICAgPEFkYXB0aXZlQ2FyZEF0dGFjaG1lbnQgYXR0YWNobWVudD17YXR0YWNobWVudH0gLz5cbiAgICAgICkgOiBhdHRhY2htZW50LmNvbnRlbnRUeXBlID09PSAnYXBwbGljYXRpb24vdm5kLm1pY3Jvc29mdC5jYXJkLmFuaW1hdGlvbicgPyAoXG4gICAgICAgIDxBbmltYXRpb25DYXJkQXR0YWNobWVudCBhdHRhY2htZW50PXthdHRhY2htZW50fSAvPlxuICAgICAgKSA6IGF0dGFjaG1lbnQuY29udGVudFR5cGUgPT09ICdhcHBsaWNhdGlvbi92bmQubWljcm9zb2Z0LmNhcmQuYXVkaW8nID8gKFxuICAgICAgICA8QXVkaW9DYXJkQXR0YWNobWVudCBhdHRhY2htZW50PXthdHRhY2htZW50fSAvPlxuICAgICAgKSA6IGF0dGFjaG1lbnQuY29udGVudFR5cGUgPT09ICdhcHBsaWNhdGlvbi92bmQubWljcm9zb2Z0LmNhcmQub2F1dGgnID8gKFxuICAgICAgICA8T0F1dGhDYXJkQXR0YWNobWVudCBhdHRhY2htZW50PXthdHRhY2htZW50fSAvPlxuICAgICAgKSA6IGF0dGFjaG1lbnQuY29udGVudFR5cGUgPT09ICdhcHBsaWNhdGlvbi92bmQubWljcm9zb2Z0LmNhcmQucmVjZWlwdCcgPyAoXG4gICAgICAgIDxSZWNlaXB0Q2FyZEF0dGFjaG1lbnQgYXR0YWNobWVudD17YXR0YWNobWVudH0gLz5cbiAgICAgICkgOiBhdHRhY2htZW50LmNvbnRlbnRUeXBlID09PSAnYXBwbGljYXRpb24vdm5kLm1pY3Jvc29mdC5jYXJkLnNpZ25pbicgPyAoXG4gICAgICAgIDxTaWduSW5DYXJkQXR0YWNobWVudCBhdHRhY2htZW50PXthdHRhY2htZW50fSAvPlxuICAgICAgKSA6IGF0dGFjaG1lbnQuY29udGVudFR5cGUgPT09ICdhcHBsaWNhdGlvbi92bmQubWljcm9zb2Z0LmNhcmQudGh1bWJuYWlsJyA/IChcbiAgICAgICAgPFRodW1ibmFpbENhcmRBdHRhY2htZW50IGF0dGFjaG1lbnQ9e2F0dGFjaG1lbnR9IC8+XG4gICAgICApIDogYXR0YWNobWVudC5jb250ZW50VHlwZSA9PT0gJ2FwcGxpY2F0aW9uL3ZuZC5taWNyb3NvZnQuY2FyZC52aWRlbycgPyAoXG4gICAgICAgIDxWaWRlb0NhcmRBdHRhY2htZW50IGF0dGFjaG1lbnQ9e2F0dGFjaG1lbnR9IC8+XG4gICAgICApIDogKFxuICAgICAgICBuZXh0KC4uLmFyZ3MpXG4gICAgICApO1xuICAgIH07XG59XG4iXSwibWFwcGluZ3MiOiI7Ozs7Ozs7OztBQUFBOztBQUlBOztBQUNBOztBQUNBOztBQUNBOztBQUNBOztBQUNBOztBQUNBOztBQUNBOztBQUNBOztBQUVlLFNBQVNBLHVDQUFULEdBQXlFO0VBQ3RGO0VBQ0EsT0FBTztJQUFBLE9BQ0wsVUFBQUMsSUFBSTtNQUFBLE9BQ0osWUFBYTtRQUFBLGtDQUFUQyxJQUFTO1VBQVRBLElBQVM7UUFBQTs7UUFDWCxJQUFTQyxVQUFULEdBQXlCRCxJQUF6QixJQUFTQyxVQUFUO1FBRUEsT0FBT0EsVUFBVSxDQUFDQyxXQUFYLEtBQTJCLHFDQUEzQixnQkFDTCw2QkFBQywyQkFBRDtVQUFvQixVQUFVLEVBQUVEO1FBQWhDLEVBREssR0FFSEEsVUFBVSxDQUFDQyxXQUFYLEtBQTJCLHlDQUEzQixnQkFDRiw2QkFBQywrQkFBRDtVQUF3QixVQUFVLEVBQUVEO1FBQXBDLEVBREUsR0FFQUEsVUFBVSxDQUFDQyxXQUFYLEtBQTJCLDBDQUEzQixnQkFDRiw2QkFBQyxnQ0FBRDtVQUF5QixVQUFVLEVBQUVEO1FBQXJDLEVBREUsR0FFQUEsVUFBVSxDQUFDQyxXQUFYLEtBQTJCLHNDQUEzQixnQkFDRiw2QkFBQyw0QkFBRDtVQUFxQixVQUFVLEVBQUVEO1FBQWpDLEVBREUsR0FFQUEsVUFBVSxDQUFDQyxXQUFYLEtBQTJCLHNDQUEzQixnQkFDRiw2QkFBQyw0QkFBRDtVQUFxQixVQUFVLEVBQUVEO1FBQWpDLEVBREUsR0FFQUEsVUFBVSxDQUFDQyxXQUFYLEtBQTJCLHdDQUEzQixnQkFDRiw2QkFBQyw4QkFBRDtVQUF1QixVQUFVLEVBQUVEO1FBQW5DLEVBREUsR0FFQUEsVUFBVSxDQUFDQyxXQUFYLEtBQTJCLHVDQUEzQixnQkFDRiw2QkFBQyw2QkFBRDtVQUFzQixVQUFVLEVBQUVEO1FBQWxDLEVBREUsR0FFQUEsVUFBVSxDQUFDQyxXQUFYLEtBQTJCLDBDQUEzQixnQkFDRiw2QkFBQyxnQ0FBRDtVQUF5QixVQUFVLEVBQUVEO1FBQXJDLEVBREUsR0FFQUEsVUFBVSxDQUFDQyxXQUFYLEtBQTJCLHNDQUEzQixnQkFDRiw2QkFBQyw0QkFBRDtVQUFxQixVQUFVLEVBQUVEO1FBQWpDLEVBREUsR0FHRkYsSUFBSSxNQUFKLFNBQVFDLElBQVIsQ0FuQkY7TUFxQkQsQ0F6Qkc7SUFBQSxDQURDO0VBQUEsQ0FBUDtBQTJCRCJ9