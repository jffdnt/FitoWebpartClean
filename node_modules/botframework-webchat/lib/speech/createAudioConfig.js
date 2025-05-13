"use strict";

var _interopRequireDefault = require("@babel/runtime/helpers/interopRequireDefault");

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.default = createAudioConfig;

var _classCallCheck2 = _interopRequireDefault(require("@babel/runtime/helpers/classCallCheck"));

var _createClass2 = _interopRequireDefault(require("@babel/runtime/helpers/createClass"));

var _assertThisInitialized2 = _interopRequireDefault(require("@babel/runtime/helpers/assertThisInitialized"));

var _inherits2 = _interopRequireDefault(require("@babel/runtime/helpers/inherits"));

var _possibleConstructorReturn2 = _interopRequireDefault(require("@babel/runtime/helpers/possibleConstructorReturn"));

var _getPrototypeOf2 = _interopRequireDefault(require("@babel/runtime/helpers/getPrototypeOf"));

var _defineProperty2 = _interopRequireDefault(require("@babel/runtime/helpers/defineProperty"));

var _microsoftCognitiveservicesSpeechSdk = require("microsoft-cognitiveservices-speech-sdk");

var _CustomAudioInputStream = _interopRequireDefault(require("./CustomAudioInputStream"));

function _createSuper(Derived) { var hasNativeReflectConstruct = _isNativeReflectConstruct(); return function _createSuperInternal() { var Super = (0, _getPrototypeOf2.default)(Derived), result; if (hasNativeReflectConstruct) { var NewTarget = (0, _getPrototypeOf2.default)(this).constructor; result = Reflect.construct(Super, arguments, NewTarget); } else { result = Super.apply(this, arguments); } return (0, _possibleConstructorReturn2.default)(this, result); }; }

function _isNativeReflectConstruct() { if (typeof Reflect === "undefined" || !Reflect.construct) return false; if (Reflect.construct.sham) return false; if (typeof Proxy === "function") return true; try { Boolean.prototype.valueOf.call(Reflect.construct(Boolean, [], function () {})); return true; } catch (e) { return false; } }

var SYMBOL_ATTACH = Symbol('attach');
var SYMBOL_TURN_OFF = Symbol('turnOff');

var CreateAudioConfigAudioInputStream = /*#__PURE__*/function (_CustomAudioInputStre) {
  (0, _inherits2.default)(CreateAudioConfigAudioInputStream, _CustomAudioInputStre);

  var _super = _createSuper(CreateAudioConfigAudioInputStream);

  function CreateAudioConfigAudioInputStream(_ref) {
    var _this;

    var attach = _ref.attach,
        debug = _ref.debug,
        turnOff = _ref.turnOff;
    (0, _classCallCheck2.default)(this, CreateAudioConfigAudioInputStream);

    if (!attach || typeof attach !== 'function') {
      throw new Error('"attach" must be a function.');
    }

    if (turnOff && typeof turnOff !== 'function') {
      throw new Error('"turnOff", if defined, must be a function.');
    }

    _this = _super.call(this, {
      debug: debug
    }); // False alarm: indexer is a constant of type Symbol.
    // eslint-disable-next-line security/detect-object-injection

    (0, _defineProperty2.default)((0, _assertThisInitialized2.default)(_this), SYMBOL_ATTACH, void 0);
    (0, _defineProperty2.default)((0, _assertThisInitialized2.default)(_this), SYMBOL_TURN_OFF, void 0);
    _this[SYMBOL_ATTACH] = attach; // False alarm: indexer is a constant of type Symbol.
    // eslint-disable-next-line security/detect-object-injection

    _this[SYMBOL_TURN_OFF] = turnOff;
    return _this;
  }

  (0, _createClass2.default)(CreateAudioConfigAudioInputStream, [{
    key: "performAttach",
    value: function performAttach(audioNodeId) {
      // False alarm: indexer is a constant of type Symbol.
      // eslint-disable-next-line security/detect-object-injection
      return this[SYMBOL_ATTACH](audioNodeId);
    }
  }, {
    key: "performTurnOff",
    value: function performTurnOff() {
      var turnOff = this[SYMBOL_TURN_OFF];
      return turnOff && turnOff();
    }
  }]);
  return CreateAudioConfigAudioInputStream;
}(_CustomAudioInputStream.default);

function createAudioConfig(options) {
  return _microsoftCognitiveservicesSpeechSdk.AudioConfig.fromStreamInput(new CreateAudioConfigAudioInputStream(options));
}
//# sourceMappingURL=data:application/json;charset=utf-8;base64,eyJ2ZXJzaW9uIjozLCJuYW1lcyI6WyJTWU1CT0xfQVRUQUNIIiwiU3ltYm9sIiwiU1lNQk9MX1RVUk5fT0ZGIiwiQ3JlYXRlQXVkaW9Db25maWdBdWRpb0lucHV0U3RyZWFtIiwiYXR0YWNoIiwiZGVidWciLCJ0dXJuT2ZmIiwiRXJyb3IiLCJhdWRpb05vZGVJZCIsIkN1c3RvbUF1ZGlvSW5wdXRTdHJlYW0iLCJjcmVhdGVBdWRpb0NvbmZpZyIsIm9wdGlvbnMiLCJBdWRpb0NvbmZpZyIsImZyb21TdHJlYW1JbnB1dCJdLCJzb3VyY2VSb290IjoiYnVuZGxlOi8vLyIsInNvdXJjZXMiOlsiLi4vLi4vc3JjL3NwZWVjaC9jcmVhdGVBdWRpb0NvbmZpZy50cyJdLCJzb3VyY2VzQ29udGVudCI6WyIvLyBUT0RPOiBbUDJdICMzOTc2IFdlIHNob3VsZCBleHBvcnQgdGhpcyB0byBhbGxvdyB3ZWIgZGV2ZWxvcGVycyB0byBicmluZyBpbiB0aGVpciBvd24gbWljcm9waG9uZS5cbi8vICAgICAgIEZvciBleGFtcGxlLCBpdCBzaG91bGQgZW5hYmxlIFJlYWN0IE5hdGl2ZSBkZXZzIHRvIGJyaW5nIGluIHRoZWlyIG1pY3JvcGhvbmUgaW1wbGVtZW50YXRpb24gYW5kIHVzZSBDb2duaXRpdmUgU2VydmljZXMgU3BlZWNoIFNlcnZpY2VzLlxuXG5pbXBvcnQgeyBBdWRpb0NvbmZpZyB9IGZyb20gJ21pY3Jvc29mdC1jb2duaXRpdmVzZXJ2aWNlcy1zcGVlY2gtc2RrJztcblxuaW1wb3J0IEN1c3RvbUF1ZGlvSW5wdXRTdHJlYW0sIHsgQXVkaW9TdHJlYW1Ob2RlLCBEZXZpY2VJbmZvLCBGb3JtYXQgfSBmcm9tICcuL0N1c3RvbUF1ZGlvSW5wdXRTdHJlYW0nO1xuXG50eXBlIEF0dGFjaEZ1bmN0aW9uID0gKGF1ZGlvTm9kZUlkOiBzdHJpbmcpID0+IFByb21pc2U8e1xuICBhdWRpb1N0cmVhbU5vZGU6IEF1ZGlvU3RyZWFtTm9kZTtcbiAgZGV2aWNlSW5mbzogRGV2aWNlSW5mbztcbiAgZm9ybWF0OiBGb3JtYXQ7XG59PjtcblxudHlwZSBUdXJuT2ZmRnVuY3Rpb24gPSAoKSA9PiBQcm9taXNlPHZvaWQ+O1xuXG5jb25zdCBTWU1CT0xfQVRUQUNIID0gU3ltYm9sKCdhdHRhY2gnKTtcbmNvbnN0IFNZTUJPTF9UVVJOX09GRiA9IFN5bWJvbCgndHVybk9mZicpO1xuXG50eXBlIENyZWF0ZUF1ZGlvQ29uZmlnT3B0aW9ucyA9IHtcbiAgLyoqIENhbGxiYWNrIGZ1bmN0aW9uIGZvciBhdHRhY2hpbmcgdGhlIGRldmljZSBieSByZXR1cm5pbmcgYW4gYXVkaW8gbm9kZS4gKi9cbiAgYXR0YWNoOiBBdHRhY2hGdW5jdGlvbjtcblxuICAvKiogYHRydWVgIHRvIGVuYWJsZSBkaWFnbm9zdGljIGluZm9ybWF0aW9uLCBvdGhlcndpc2UsIGBmYWxzZWAuICovXG4gIGRlYnVnPzogdHJ1ZTtcblxuICAvKipcbiAgICogQ2FsbGJhY2sgZnVuY3Rpb24gZm9yIHR1cm5pbmcgb2ZmIHRoZSBkZXZpY2UgYmVmb3JlIGRldGFjaGluZyBpdHMgbm9kZSBmcm9tIGFuIGF1ZGlvIGdyYXBoLlxuICAgKlxuICAgKiBOb3RlOiB0aGlzIGlzIG5vdCBjYWxsZWQgZm9yIERpcmVjdCBMaW5lIFNwZWVjaC5cbiAgICovXG4gIHR1cm5PZmY/OiBUdXJuT2ZmRnVuY3Rpb247XG59O1xuXG5jbGFzcyBDcmVhdGVBdWRpb0NvbmZpZ0F1ZGlvSW5wdXRTdHJlYW0gZXh0ZW5kcyBDdXN0b21BdWRpb0lucHV0U3RyZWFtIHtcbiAgY29uc3RydWN0b3IoeyBhdHRhY2gsIGRlYnVnLCB0dXJuT2ZmIH06IENyZWF0ZUF1ZGlvQ29uZmlnT3B0aW9ucykge1xuICAgIGlmICghYXR0YWNoIHx8IHR5cGVvZiBhdHRhY2ggIT09ICdmdW5jdGlvbicpIHtcbiAgICAgIHRocm93IG5ldyBFcnJvcignXCJhdHRhY2hcIiBtdXN0IGJlIGEgZnVuY3Rpb24uJyk7XG4gICAgfVxuXG4gICAgaWYgKHR1cm5PZmYgJiYgdHlwZW9mIHR1cm5PZmYgIT09ICdmdW5jdGlvbicpIHtcbiAgICAgIHRocm93IG5ldyBFcnJvcignXCJ0dXJuT2ZmXCIsIGlmIGRlZmluZWQsIG11c3QgYmUgYSBmdW5jdGlvbi4nKTtcbiAgICB9XG5cbiAgICBzdXBlcih7IGRlYnVnIH0pO1xuXG4gICAgLy8gRmFsc2UgYWxhcm06IGluZGV4ZXIgaXMgYSBjb25zdGFudCBvZiB0eXBlIFN5bWJvbC5cbiAgICAvLyBlc2xpbnQtZGlzYWJsZS1uZXh0LWxpbmUgc2VjdXJpdHkvZGV0ZWN0LW9iamVjdC1pbmplY3Rpb25cbiAgICB0aGlzW1NZTUJPTF9BVFRBQ0hdID0gYXR0YWNoO1xuXG4gICAgLy8gRmFsc2UgYWxhcm06IGluZGV4ZXIgaXMgYSBjb25zdGFudCBvZiB0eXBlIFN5bWJvbC5cbiAgICAvLyBlc2xpbnQtZGlzYWJsZS1uZXh0LWxpbmUgc2VjdXJpdHkvZGV0ZWN0LW9iamVjdC1pbmplY3Rpb25cbiAgICB0aGlzW1NZTUJPTF9UVVJOX09GRl0gPSB0dXJuT2ZmO1xuICB9XG5cbiAgW1NZTUJPTF9BVFRBQ0hdOiBBdHRhY2hGdW5jdGlvbjtcbiAgW1NZTUJPTF9UVVJOX09GRl06IFR1cm5PZmZGdW5jdGlvbjtcblxuICBwcm90ZWN0ZWQgcGVyZm9ybUF0dGFjaChhdWRpb05vZGVJZDogc3RyaW5nKTogUHJvbWlzZTx7XG4gICAgYXVkaW9TdHJlYW1Ob2RlOiBBdWRpb1N0cmVhbU5vZGU7XG4gICAgZGV2aWNlSW5mbzogRGV2aWNlSW5mbztcbiAgICBmb3JtYXQ6IEZvcm1hdDtcbiAgfT4ge1xuICAgIC8vIEZhbHNlIGFsYXJtOiBpbmRleGVyIGlzIGEgY29uc3RhbnQgb2YgdHlwZSBTeW1ib2wuXG4gICAgLy8gZXNsaW50LWRpc2FibGUtbmV4dC1saW5lIHNlY3VyaXR5L2RldGVjdC1vYmplY3QtaW5qZWN0aW9uXG4gICAgcmV0dXJuIHRoaXNbU1lNQk9MX0FUVEFDSF0oYXVkaW9Ob2RlSWQpO1xuICB9XG5cbiAgcHJvdGVjdGVkIHBlcmZvcm1UdXJuT2ZmKCk6IFByb21pc2U8dm9pZD4ge1xuICAgIGNvbnN0IHsgW1NZTUJPTF9UVVJOX09GRl06IHR1cm5PZmYgfSA9IHRoaXM7XG5cbiAgICByZXR1cm4gdHVybk9mZiAmJiB0dXJuT2ZmKCk7XG4gIH1cbn1cblxuZXhwb3J0IGRlZmF1bHQgZnVuY3Rpb24gY3JlYXRlQXVkaW9Db25maWcob3B0aW9uczogQ3JlYXRlQXVkaW9Db25maWdPcHRpb25zKSB7XG4gIHJldHVybiBBdWRpb0NvbmZpZy5mcm9tU3RyZWFtSW5wdXQobmV3IENyZWF0ZUF1ZGlvQ29uZmlnQXVkaW9JbnB1dFN0cmVhbShvcHRpb25zKSk7XG59XG4iXSwibWFwcGluZ3MiOiI7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7O0FBR0E7O0FBRUE7Ozs7OztBQVVBLElBQU1BLGFBQWEsR0FBR0MsTUFBTSxDQUFDLFFBQUQsQ0FBNUI7QUFDQSxJQUFNQyxlQUFlLEdBQUdELE1BQU0sQ0FBQyxTQUFELENBQTlCOztJQWlCTUUsaUM7Ozs7O0VBQ0osaURBQWtFO0lBQUE7O0lBQUEsSUFBcERDLE1BQW9ELFFBQXBEQSxNQUFvRDtJQUFBLElBQTVDQyxLQUE0QyxRQUE1Q0EsS0FBNEM7SUFBQSxJQUFyQ0MsT0FBcUMsUUFBckNBLE9BQXFDO0lBQUE7O0lBQ2hFLElBQUksQ0FBQ0YsTUFBRCxJQUFXLE9BQU9BLE1BQVAsS0FBa0IsVUFBakMsRUFBNkM7TUFDM0MsTUFBTSxJQUFJRyxLQUFKLENBQVUsOEJBQVYsQ0FBTjtJQUNEOztJQUVELElBQUlELE9BQU8sSUFBSSxPQUFPQSxPQUFQLEtBQW1CLFVBQWxDLEVBQThDO01BQzVDLE1BQU0sSUFBSUMsS0FBSixDQUFVLDRDQUFWLENBQU47SUFDRDs7SUFFRCwwQkFBTTtNQUFFRixLQUFLLEVBQUxBO0lBQUYsQ0FBTixFQVRnRSxDQVdoRTtJQUNBOztJQVpnRSwyRUFvQmpFTCxhQXBCaUU7SUFBQSwyRUFxQmpFRSxlQXJCaUU7SUFhaEUsTUFBS0YsYUFBTCxJQUFzQkksTUFBdEIsQ0FiZ0UsQ0FlaEU7SUFDQTs7SUFDQSxNQUFLRixlQUFMLElBQXdCSSxPQUF4QjtJQWpCZ0U7RUFrQmpFOzs7O1dBS0QsdUJBQXdCRSxXQUF4QixFQUlHO01BQ0Q7TUFDQTtNQUNBLE9BQU8sS0FBS1IsYUFBTCxFQUFvQlEsV0FBcEIsQ0FBUDtJQUNEOzs7V0FFRCwwQkFBMEM7TUFDeEMsSUFBMkJGLE9BQTNCLEdBQXVDLElBQXZDLENBQVNKLGVBQVQ7TUFFQSxPQUFPSSxPQUFPLElBQUlBLE9BQU8sRUFBekI7SUFDRDs7O0VBdEM2Q0csK0I7O0FBeUNqQyxTQUFTQyxpQkFBVCxDQUEyQkMsT0FBM0IsRUFBOEQ7RUFDM0UsT0FBT0MsZ0RBQUEsQ0FBWUMsZUFBWixDQUE0QixJQUFJVixpQ0FBSixDQUFzQ1EsT0FBdEMsQ0FBNUIsQ0FBUDtBQUNEIn0=