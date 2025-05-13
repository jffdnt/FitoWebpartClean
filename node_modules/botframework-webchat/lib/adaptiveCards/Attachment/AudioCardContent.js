"use strict";

var _interopRequireDefault = require("@babel/runtime/helpers/interopRequireDefault");

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.default = void 0;

var _slicedToArray2 = _interopRequireDefault(require("@babel/runtime/helpers/slicedToArray"));

var _botframeworkWebchatComponent = require("botframework-webchat-component");

var _propTypes = _interopRequireDefault(require("prop-types"));

var _react = _interopRequireDefault(require("react"));

var _CommonCard = _interopRequireDefault(require("./CommonCard"));

/* eslint react/no-array-index-key: "off" */
var AudioContent = _botframeworkWebchatComponent.Components.AudioContent;
var useStyleSet = _botframeworkWebchatComponent.hooks.useStyleSet;

var AudioCardContent = function AudioCardContent(_ref) {
  var actionPerformedClassName = _ref.actionPerformedClassName,
      content = _ref.content,
      disabled = _ref.disabled;

  var _useStyleSet = useStyleSet(),
      _useStyleSet2 = (0, _slicedToArray2.default)(_useStyleSet, 1),
      audioCardAttachmentStyleSet = _useStyleSet2[0].audioCardAttachment;

  var _content$autostart = content.autostart,
      autostart = _content$autostart === void 0 ? false : _content$autostart,
      _content$autoloop = content.autoloop,
      autoloop = _content$autoloop === void 0 ? false : _content$autoloop,
      _content$image = content.image;
  _content$image = _content$image === void 0 ? {} : _content$image;
  var _content$image$url = _content$image.url,
      imageURL = _content$image$url === void 0 ? '' : _content$image$url,
      _content$media = content.media,
      media = _content$media === void 0 ? [] : _content$media;
  return /*#__PURE__*/_react.default.createElement("div", {
    className: audioCardAttachmentStyleSet
  }, /*#__PURE__*/_react.default.createElement("ul", {
    className: "media-list"
  }, media.map(function (_ref2, index) {
    var url = _ref2.url;
    return /*#__PURE__*/_react.default.createElement("li", {
      key: index
    }, /*#__PURE__*/_react.default.createElement(AudioContent, {
      autoPlay: autostart,
      loop: autoloop,
      poster: imageURL,
      src: url
    }));
  })), /*#__PURE__*/_react.default.createElement(_CommonCard.default, {
    actionPerformedClassName: actionPerformedClassName,
    content: content,
    disabled: disabled
  }));
};

AudioCardContent.defaultProps = {
  actionPerformedClassName: '',
  disabled: undefined
};
AudioCardContent.propTypes = {
  actionPerformedClassName: _propTypes.default.string,
  // PropTypes cannot fully capture TypeScript types.
  // @ts-ignore
  content: _propTypes.default.shape({
    autostart: _propTypes.default.bool,
    autoloop: _propTypes.default.bool,
    image: _propTypes.default.shape({
      url: _propTypes.default.string.isRequired
    }),
    media: _propTypes.default.arrayOf(_propTypes.default.shape({
      url: _propTypes.default.string.isRequired
    }).isRequired).isRequired
  }).isRequired,
  disabled: _propTypes.default.bool
};
var _default = AudioCardContent;
exports.default = _default;
//# sourceMappingURL=data:application/json;charset=utf-8;base64,eyJ2ZXJzaW9uIjozLCJuYW1lcyI6WyJBdWRpb0NvbnRlbnQiLCJDb21wb25lbnRzIiwidXNlU3R5bGVTZXQiLCJob29rcyIsIkF1ZGlvQ2FyZENvbnRlbnQiLCJhY3Rpb25QZXJmb3JtZWRDbGFzc05hbWUiLCJjb250ZW50IiwiZGlzYWJsZWQiLCJhdWRpb0NhcmRBdHRhY2htZW50U3R5bGVTZXQiLCJhdWRpb0NhcmRBdHRhY2htZW50IiwiYXV0b3N0YXJ0IiwiYXV0b2xvb3AiLCJpbWFnZSIsInVybCIsImltYWdlVVJMIiwibWVkaWEiLCJtYXAiLCJpbmRleCIsImRlZmF1bHRQcm9wcyIsInVuZGVmaW5lZCIsInByb3BUeXBlcyIsIlByb3BUeXBlcyIsInN0cmluZyIsInNoYXBlIiwiYm9vbCIsImlzUmVxdWlyZWQiLCJhcnJheU9mIl0sInNvdXJjZVJvb3QiOiJidW5kbGU6Ly8vIiwic291cmNlcyI6WyIuLi8uLi8uLi9zcmMvYWRhcHRpdmVDYXJkcy9BdHRhY2htZW50L0F1ZGlvQ2FyZENvbnRlbnQudHN4Il0sInNvdXJjZXNDb250ZW50IjpbIi8qIGVzbGludCByZWFjdC9uby1hcnJheS1pbmRleC1rZXk6IFwib2ZmXCIgKi9cblxuaW1wb3J0IHsgQ29tcG9uZW50cywgaG9va3MgfSBmcm9tICdib3RmcmFtZXdvcmstd2ViY2hhdC1jb21wb25lbnQnO1xuaW1wb3J0IFByb3BUeXBlcyBmcm9tICdwcm9wLXR5cGVzJztcbmltcG9ydCBSZWFjdCwgeyBGQyB9IGZyb20gJ3JlYWN0JztcbmltcG9ydCB0eXBlIHsgRGlyZWN0TGluZUF1ZGlvQ2FyZCB9IGZyb20gJ2JvdGZyYW1ld29yay13ZWJjaGF0LWNvcmUnO1xuXG5pbXBvcnQgQ29tbW9uQ2FyZCBmcm9tICcuL0NvbW1vbkNhcmQnO1xuXG5jb25zdCB7IEF1ZGlvQ29udGVudCB9ID0gQ29tcG9uZW50cztcbmNvbnN0IHsgdXNlU3R5bGVTZXQgfSA9IGhvb2tzO1xuXG50eXBlIEF1ZGlvQ2FyZENvbnRlbnRQcm9wcyA9IHtcbiAgYWN0aW9uUGVyZm9ybWVkQ2xhc3NOYW1lPzogc3RyaW5nO1xuICBjb250ZW50OiBEaXJlY3RMaW5lQXVkaW9DYXJkO1xuICBkaXNhYmxlZD86IGJvb2xlYW47XG59O1xuXG5jb25zdCBBdWRpb0NhcmRDb250ZW50OiBGQzxBdWRpb0NhcmRDb250ZW50UHJvcHM+ID0gKHsgYWN0aW9uUGVyZm9ybWVkQ2xhc3NOYW1lLCBjb250ZW50LCBkaXNhYmxlZCB9KSA9PiB7XG4gIGNvbnN0IFt7IGF1ZGlvQ2FyZEF0dGFjaG1lbnQ6IGF1ZGlvQ2FyZEF0dGFjaG1lbnRTdHlsZVNldCB9XSA9IHVzZVN0eWxlU2V0KCk7XG4gIGNvbnN0IHsgYXV0b3N0YXJ0ID0gZmFsc2UsIGF1dG9sb29wID0gZmFsc2UsIGltYWdlOiB7IHVybDogaW1hZ2VVUkwgPSAnJyB9ID0ge30sIG1lZGlhID0gW10gfSA9IGNvbnRlbnQ7XG5cbiAgcmV0dXJuIChcbiAgICA8ZGl2IGNsYXNzTmFtZT17YXVkaW9DYXJkQXR0YWNobWVudFN0eWxlU2V0fT5cbiAgICAgIDx1bCBjbGFzc05hbWU9XCJtZWRpYS1saXN0XCI+XG4gICAgICAgIHttZWRpYS5tYXAoKHsgdXJsIH0sIGluZGV4KSA9PiAoXG4gICAgICAgICAgPGxpIGtleT17aW5kZXh9PlxuICAgICAgICAgICAgPEF1ZGlvQ29udGVudCBhdXRvUGxheT17YXV0b3N0YXJ0fSBsb29wPXthdXRvbG9vcH0gcG9zdGVyPXtpbWFnZVVSTH0gc3JjPXt1cmx9IC8+XG4gICAgICAgICAgPC9saT5cbiAgICAgICAgKSl9XG4gICAgICA8L3VsPlxuICAgICAgPENvbW1vbkNhcmQgYWN0aW9uUGVyZm9ybWVkQ2xhc3NOYW1lPXthY3Rpb25QZXJmb3JtZWRDbGFzc05hbWV9IGNvbnRlbnQ9e2NvbnRlbnR9IGRpc2FibGVkPXtkaXNhYmxlZH0gLz5cbiAgICA8L2Rpdj5cbiAgKTtcbn07XG5cbkF1ZGlvQ2FyZENvbnRlbnQuZGVmYXVsdFByb3BzID0ge1xuICBhY3Rpb25QZXJmb3JtZWRDbGFzc05hbWU6ICcnLFxuICBkaXNhYmxlZDogdW5kZWZpbmVkXG59O1xuXG5BdWRpb0NhcmRDb250ZW50LnByb3BUeXBlcyA9IHtcbiAgYWN0aW9uUGVyZm9ybWVkQ2xhc3NOYW1lOiBQcm9wVHlwZXMuc3RyaW5nLFxuICAvLyBQcm9wVHlwZXMgY2Fubm90IGZ1bGx5IGNhcHR1cmUgVHlwZVNjcmlwdCB0eXBlcy5cbiAgLy8gQHRzLWlnbm9yZVxuICBjb250ZW50OiBQcm9wVHlwZXMuc2hhcGUoe1xuICAgIGF1dG9zdGFydDogUHJvcFR5cGVzLmJvb2wsXG4gICAgYXV0b2xvb3A6IFByb3BUeXBlcy5ib29sLFxuICAgIGltYWdlOiBQcm9wVHlwZXMuc2hhcGUoe1xuICAgICAgdXJsOiBQcm9wVHlwZXMuc3RyaW5nLmlzUmVxdWlyZWRcbiAgICB9KSxcbiAgICBtZWRpYTogUHJvcFR5cGVzLmFycmF5T2YoXG4gICAgICBQcm9wVHlwZXMuc2hhcGUoe1xuICAgICAgICB1cmw6IFByb3BUeXBlcy5zdHJpbmcuaXNSZXF1aXJlZFxuICAgICAgfSkuaXNSZXF1aXJlZFxuICAgICkuaXNSZXF1aXJlZFxuICB9KS5pc1JlcXVpcmVkLFxuICBkaXNhYmxlZDogUHJvcFR5cGVzLmJvb2xcbn07XG5cbmV4cG9ydCBkZWZhdWx0IEF1ZGlvQ2FyZENvbnRlbnQ7XG4iXSwibWFwcGluZ3MiOiI7Ozs7Ozs7Ozs7O0FBRUE7O0FBQ0E7O0FBQ0E7O0FBR0E7O0FBUEE7QUFTQSxJQUFRQSxZQUFSLEdBQXlCQyx3Q0FBekIsQ0FBUUQsWUFBUjtBQUNBLElBQVFFLFdBQVIsR0FBd0JDLG1DQUF4QixDQUFRRCxXQUFSOztBQVFBLElBQU1FLGdCQUEyQyxHQUFHLFNBQTlDQSxnQkFBOEMsT0FBcUQ7RUFBQSxJQUFsREMsd0JBQWtELFFBQWxEQSx3QkFBa0Q7RUFBQSxJQUF4QkMsT0FBd0IsUUFBeEJBLE9BQXdCO0VBQUEsSUFBZkMsUUFBZSxRQUFmQSxRQUFlOztFQUN2RyxtQkFBK0RMLFdBQVcsRUFBMUU7RUFBQTtFQUFBLElBQThCTSwyQkFBOUIsb0JBQVNDLG1CQUFUOztFQUNBLHlCQUFnR0gsT0FBaEcsQ0FBUUksU0FBUjtFQUFBLElBQVFBLFNBQVIsbUNBQW9CLEtBQXBCO0VBQUEsd0JBQWdHSixPQUFoRyxDQUEyQkssUUFBM0I7RUFBQSxJQUEyQkEsUUFBM0Isa0NBQXNDLEtBQXRDO0VBQUEscUJBQWdHTCxPQUFoRyxDQUE2Q00sS0FBN0M7RUFBQSw2Q0FBNkUsRUFBN0U7RUFBQSx3Q0FBc0RDLEdBQXREO0VBQUEsSUFBMkRDLFFBQTNELG1DQUFzRSxFQUF0RTtFQUFBLHFCQUFnR1IsT0FBaEcsQ0FBaUZTLEtBQWpGO0VBQUEsSUFBaUZBLEtBQWpGLCtCQUF5RixFQUF6RjtFQUVBLG9CQUNFO0lBQUssU0FBUyxFQUFFUDtFQUFoQixnQkFDRTtJQUFJLFNBQVMsRUFBQztFQUFkLEdBQ0dPLEtBQUssQ0FBQ0MsR0FBTixDQUFVLGlCQUFVQyxLQUFWO0lBQUEsSUFBR0osR0FBSCxTQUFHQSxHQUFIO0lBQUEsb0JBQ1Q7TUFBSSxHQUFHLEVBQUVJO0lBQVQsZ0JBQ0UsNkJBQUMsWUFBRDtNQUFjLFFBQVEsRUFBRVAsU0FBeEI7TUFBbUMsSUFBSSxFQUFFQyxRQUF6QztNQUFtRCxNQUFNLEVBQUVHLFFBQTNEO01BQXFFLEdBQUcsRUFBRUQ7SUFBMUUsRUFERixDQURTO0VBQUEsQ0FBVixDQURILENBREYsZUFRRSw2QkFBQyxtQkFBRDtJQUFZLHdCQUF3QixFQUFFUix3QkFBdEM7SUFBZ0UsT0FBTyxFQUFFQyxPQUF6RTtJQUFrRixRQUFRLEVBQUVDO0VBQTVGLEVBUkYsQ0FERjtBQVlELENBaEJEOztBQWtCQUgsZ0JBQWdCLENBQUNjLFlBQWpCLEdBQWdDO0VBQzlCYix3QkFBd0IsRUFBRSxFQURJO0VBRTlCRSxRQUFRLEVBQUVZO0FBRm9CLENBQWhDO0FBS0FmLGdCQUFnQixDQUFDZ0IsU0FBakIsR0FBNkI7RUFDM0JmLHdCQUF3QixFQUFFZ0Isa0JBQUEsQ0FBVUMsTUFEVDtFQUUzQjtFQUNBO0VBQ0FoQixPQUFPLEVBQUVlLGtCQUFBLENBQVVFLEtBQVYsQ0FBZ0I7SUFDdkJiLFNBQVMsRUFBRVcsa0JBQUEsQ0FBVUcsSUFERTtJQUV2QmIsUUFBUSxFQUFFVSxrQkFBQSxDQUFVRyxJQUZHO0lBR3ZCWixLQUFLLEVBQUVTLGtCQUFBLENBQVVFLEtBQVYsQ0FBZ0I7TUFDckJWLEdBQUcsRUFBRVEsa0JBQUEsQ0FBVUMsTUFBVixDQUFpQkc7SUFERCxDQUFoQixDQUhnQjtJQU12QlYsS0FBSyxFQUFFTSxrQkFBQSxDQUFVSyxPQUFWLENBQ0xMLGtCQUFBLENBQVVFLEtBQVYsQ0FBZ0I7TUFDZFYsR0FBRyxFQUFFUSxrQkFBQSxDQUFVQyxNQUFWLENBQWlCRztJQURSLENBQWhCLEVBRUdBLFVBSEUsRUFJTEE7RUFWcUIsQ0FBaEIsRUFXTkEsVUFmd0I7RUFnQjNCbEIsUUFBUSxFQUFFYyxrQkFBQSxDQUFVRztBQWhCTyxDQUE3QjtlQW1CZXBCLGdCIn0=