"use strict";

var _interopRequireDefault = require("@babel/runtime/helpers/interopRequireDefault");

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.default = createAdaptiveCardsStyleSet;

var _defineProperty2 = _interopRequireDefault(require("@babel/runtime/helpers/defineProperty"));

var _botframeworkWebchatApi = require("botframework-webchat-api");

var _AdaptiveCardRenderer = _interopRequireDefault(require("./StyleSet/AdaptiveCardRenderer"));

var _AnimationCardAttachment = _interopRequireDefault(require("./StyleSet/AnimationCardAttachment"));

var _AudioCardAttachment = _interopRequireDefault(require("./StyleSet/AudioCardAttachment"));

var _normalizeStyleOptions = _interopRequireDefault(require("../normalizeStyleOptions"));

function ownKeys(object, enumerableOnly) { var keys = Object.keys(object); if (Object.getOwnPropertySymbols) { var symbols = Object.getOwnPropertySymbols(object); enumerableOnly && (symbols = symbols.filter(function (sym) { return Object.getOwnPropertyDescriptor(object, sym).enumerable; })), keys.push.apply(keys, symbols); } return keys; }

function _objectSpread(target) { for (var i = 1; i < arguments.length; i++) { var source = null != arguments[i] ? arguments[i] : {}; i % 2 ? ownKeys(Object(source), !0).forEach(function (key) { (0, _defineProperty2.default)(target, key, source[key]); }) : Object.getOwnPropertyDescriptors ? Object.defineProperties(target, Object.getOwnPropertyDescriptors(source)) : ownKeys(Object(source)).forEach(function (key) { Object.defineProperty(target, key, Object.getOwnPropertyDescriptor(source, key)); }); } return target; }

// TODO: [P4] We should add a notice for people who want to use "styleSet" instead of "styleOptions".
//       "styleSet" is actually CSS stylesheet and it is based on the DOM tree.
//       DOM tree may change from time to time, thus, maintaining "styleSet" becomes a constant effort.
function createAdaptiveCardsStyleSet(options) {
  var strictOptions = _objectSpread(_objectSpread({}, (0, _botframeworkWebchatApi.normalizeStyleOptions)(options)), (0, _normalizeStyleOptions.default)(options));

  return {
    adaptiveCardRenderer: (0, _AdaptiveCardRenderer.default)(strictOptions),
    animationCardAttachment: (0, _AnimationCardAttachment.default)(),
    audioCardAttachment: (0, _AudioCardAttachment.default)(strictOptions)
  };
}
//# sourceMappingURL=data:application/json;charset=utf-8;base64,eyJ2ZXJzaW9uIjozLCJuYW1lcyI6WyJjcmVhdGVBZGFwdGl2ZUNhcmRzU3R5bGVTZXQiLCJvcHRpb25zIiwic3RyaWN0T3B0aW9ucyIsIm5vcm1hbGl6ZVN0eWxlT3B0aW9ucyIsIm5vcm1hbGl6ZUFkYXB0aXZlQ2FyZHNTdHlsZU9wdGlvbnMiLCJhZGFwdGl2ZUNhcmRSZW5kZXJlciIsImNyZWF0ZUFkYXB0aXZlQ2FyZFJlbmRlcmVyU3R5bGUiLCJhbmltYXRpb25DYXJkQXR0YWNobWVudCIsImNyZWF0ZUFuaW1hdGlvbkNhcmRBdHRhY2htZW50U3R5bGUiLCJhdWRpb0NhcmRBdHRhY2htZW50IiwiY3JlYXRlQXVkaW9DYXJkQXR0YWNobWVudFN0eWxlIl0sInNvdXJjZVJvb3QiOiJidW5kbGU6Ly8vIiwic291cmNlcyI6WyIuLi8uLi8uLi9zcmMvYWRhcHRpdmVDYXJkcy9TdHlsZXMvY3JlYXRlQWRhcHRpdmVDYXJkc1N0eWxlU2V0LnRzIl0sInNvdXJjZXNDb250ZW50IjpbImltcG9ydCB7IG5vcm1hbGl6ZVN0eWxlT3B0aW9ucywgU3RyaWN0U3R5bGVPcHRpb25zLCBTdHlsZU9wdGlvbnMgfSBmcm9tICdib3RmcmFtZXdvcmstd2ViY2hhdC1hcGknO1xuXG5pbXBvcnQgQWRhcHRpdmVDYXJkc1N0eWxlT3B0aW9ucywgeyBTdHJpY3RBZGFwdGl2ZUNhcmRzU3R5bGVPcHRpb25zIH0gZnJvbSAnLi4vQWRhcHRpdmVDYXJkc1N0eWxlT3B0aW9ucyc7XG5pbXBvcnQgQWRhcHRpdmVDYXJkc1N0eWxlU2V0IGZyb20gJy4uL0FkYXB0aXZlQ2FyZHNTdHlsZVNldCc7XG5pbXBvcnQgY3JlYXRlQWRhcHRpdmVDYXJkUmVuZGVyZXJTdHlsZSBmcm9tICcuL1N0eWxlU2V0L0FkYXB0aXZlQ2FyZFJlbmRlcmVyJztcbmltcG9ydCBjcmVhdGVBbmltYXRpb25DYXJkQXR0YWNobWVudFN0eWxlIGZyb20gJy4vU3R5bGVTZXQvQW5pbWF0aW9uQ2FyZEF0dGFjaG1lbnQnO1xuaW1wb3J0IGNyZWF0ZUF1ZGlvQ2FyZEF0dGFjaG1lbnRTdHlsZSBmcm9tICcuL1N0eWxlU2V0L0F1ZGlvQ2FyZEF0dGFjaG1lbnQnO1xuaW1wb3J0IG5vcm1hbGl6ZUFkYXB0aXZlQ2FyZHNTdHlsZU9wdGlvbnMgZnJvbSAnLi4vbm9ybWFsaXplU3R5bGVPcHRpb25zJztcblxuLy8gVE9ETzogW1A0XSBXZSBzaG91bGQgYWRkIGEgbm90aWNlIGZvciBwZW9wbGUgd2hvIHdhbnQgdG8gdXNlIFwic3R5bGVTZXRcIiBpbnN0ZWFkIG9mIFwic3R5bGVPcHRpb25zXCIuXG4vLyAgICAgICBcInN0eWxlU2V0XCIgaXMgYWN0dWFsbHkgQ1NTIHN0eWxlc2hlZXQgYW5kIGl0IGlzIGJhc2VkIG9uIHRoZSBET00gdHJlZS5cbi8vICAgICAgIERPTSB0cmVlIG1heSBjaGFuZ2UgZnJvbSB0aW1lIHRvIHRpbWUsIHRodXMsIG1haW50YWluaW5nIFwic3R5bGVTZXRcIiBiZWNvbWVzIGEgY29uc3RhbnQgZWZmb3J0LlxuXG5leHBvcnQgZGVmYXVsdCBmdW5jdGlvbiBjcmVhdGVBZGFwdGl2ZUNhcmRzU3R5bGVTZXQoXG4gIG9wdGlvbnM6IFN0eWxlT3B0aW9ucyAmIEFkYXB0aXZlQ2FyZHNTdHlsZU9wdGlvbnNcbik6IEFkYXB0aXZlQ2FyZHNTdHlsZVNldCB7XG4gIGNvbnN0IHN0cmljdE9wdGlvbnM6IFN0cmljdFN0eWxlT3B0aW9ucyAmIFN0cmljdEFkYXB0aXZlQ2FyZHNTdHlsZU9wdGlvbnMgPSB7XG4gICAgLi4ubm9ybWFsaXplU3R5bGVPcHRpb25zKG9wdGlvbnMpLFxuICAgIC4uLm5vcm1hbGl6ZUFkYXB0aXZlQ2FyZHNTdHlsZU9wdGlvbnMob3B0aW9ucylcbiAgfTtcblxuICByZXR1cm4ge1xuICAgIGFkYXB0aXZlQ2FyZFJlbmRlcmVyOiBjcmVhdGVBZGFwdGl2ZUNhcmRSZW5kZXJlclN0eWxlKHN0cmljdE9wdGlvbnMpLFxuICAgIGFuaW1hdGlvbkNhcmRBdHRhY2htZW50OiBjcmVhdGVBbmltYXRpb25DYXJkQXR0YWNobWVudFN0eWxlKCksXG4gICAgYXVkaW9DYXJkQXR0YWNobWVudDogY3JlYXRlQXVkaW9DYXJkQXR0YWNobWVudFN0eWxlKHN0cmljdE9wdGlvbnMpXG4gIH07XG59XG4iXSwibWFwcGluZ3MiOiI7Ozs7Ozs7Ozs7O0FBQUE7O0FBSUE7O0FBQ0E7O0FBQ0E7O0FBQ0E7Ozs7OztBQUVBO0FBQ0E7QUFDQTtBQUVlLFNBQVNBLDJCQUFULENBQ2JDLE9BRGEsRUFFVTtFQUN2QixJQUFNQyxhQUFtRSxtQ0FDcEUsSUFBQUMsNkNBQUEsRUFBc0JGLE9BQXRCLENBRG9FLEdBRXBFLElBQUFHLDhCQUFBLEVBQW1DSCxPQUFuQyxDQUZvRSxDQUF6RTs7RUFLQSxPQUFPO0lBQ0xJLG9CQUFvQixFQUFFLElBQUFDLDZCQUFBLEVBQWdDSixhQUFoQyxDQURqQjtJQUVMSyx1QkFBdUIsRUFBRSxJQUFBQyxnQ0FBQSxHQUZwQjtJQUdMQyxtQkFBbUIsRUFBRSxJQUFBQyw0QkFBQSxFQUErQlIsYUFBL0I7RUFIaEIsQ0FBUDtBQUtEIn0=