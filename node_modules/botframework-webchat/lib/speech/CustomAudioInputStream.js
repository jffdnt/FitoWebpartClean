"use strict";

var _interopRequireDefault = require("@babel/runtime/helpers/interopRequireDefault");

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.default = void 0;

var _regenerator = _interopRequireDefault(require("@babel/runtime/regenerator"));

var _slicedToArray2 = _interopRequireDefault(require("@babel/runtime/helpers/slicedToArray"));

var _asyncToGenerator2 = _interopRequireDefault(require("@babel/runtime/helpers/asyncToGenerator"));

var _classCallCheck2 = _interopRequireDefault(require("@babel/runtime/helpers/classCallCheck"));

var _createClass2 = _interopRequireDefault(require("@babel/runtime/helpers/createClass"));

var _assertThisInitialized2 = _interopRequireDefault(require("@babel/runtime/helpers/assertThisInitialized"));

var _inherits2 = _interopRequireDefault(require("@babel/runtime/helpers/inherits"));

var _possibleConstructorReturn2 = _interopRequireDefault(require("@babel/runtime/helpers/possibleConstructorReturn"));

var _getPrototypeOf2 = _interopRequireDefault(require("@babel/runtime/helpers/getPrototypeOf"));

var _defineProperty2 = _interopRequireDefault(require("@babel/runtime/helpers/defineProperty"));

var _microsoftCognitiveservicesSpeechSdk = require("microsoft-cognitiveservices-speech-sdk");

var _Exports = require("microsoft-cognitiveservices-speech-sdk/distrib/lib/src/common/Exports");

var _AudioStreamFormat = require("microsoft-cognitiveservices-speech-sdk/distrib/lib/src/sdk/Audio/AudioStreamFormat");

var _Exports2 = require("microsoft-cognitiveservices-speech-sdk/distrib/lib/src/common.speech/Exports");

var _botframeworkWebchatCore = require("botframework-webchat-core");

var _uuid = require("uuid");

var _pDeferEs = _interopRequireDefault(require("p-defer-es5"));

function _createSuper(Derived) { var hasNativeReflectConstruct = _isNativeReflectConstruct(); return function _createSuperInternal() { var Super = (0, _getPrototypeOf2.default)(Derived), result; if (hasNativeReflectConstruct) { var NewTarget = (0, _getPrototypeOf2.default)(this).constructor; result = Reflect.construct(Super, arguments, NewTarget); } else { result = Super.apply(this, arguments); } return (0, _possibleConstructorReturn2.default)(this, result); }; }

function _isNativeReflectConstruct() { if (typeof Reflect === "undefined" || !Reflect.construct) return false; if (Reflect.construct.sham) return false; if (typeof Proxy === "function") return true; try { Boolean.prototype.valueOf.call(Reflect.construct(Boolean, [], function () {})); return true; } catch (e) { return false; } }

var SYMBOL_DEVICE_INFO_DEFERRED = Symbol('deviceInfoDeferred');
var SYMBOL_EVENTS = Symbol('events');
var SYMBOL_FORMAT_DEFERRED = Symbol('formatDeferred');
var SYMBOL_OPTIONS = Symbol('options'); // Speech SDK quirks: Only 2 lifecycle functions are actually used.
//                    They are: attach() and turnOff().
//                    Others are not used, including: blob(), close(), detach(), turnOn().

var CustomAudioInputStream = /*#__PURE__*/function (_AudioInputStream) {
  (0, _inherits2.default)(CustomAudioInputStream, _AudioInputStream);

  var _super = _createSuper(CustomAudioInputStream);

  function CustomAudioInputStream() {
    var _this;

    var options = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : {};
    (0, _classCallCheck2.default)(this, CustomAudioInputStream);
    _this = _super.call(this);
    (0, _defineProperty2.default)((0, _assertThisInitialized2.default)(_this), SYMBOL_DEVICE_INFO_DEFERRED, void 0);
    (0, _defineProperty2.default)((0, _assertThisInitialized2.default)(_this), SYMBOL_EVENTS, void 0);
    (0, _defineProperty2.default)((0, _assertThisInitialized2.default)(_this), SYMBOL_FORMAT_DEFERRED, void 0);
    (0, _defineProperty2.default)((0, _assertThisInitialized2.default)(_this), SYMBOL_OPTIONS, void 0);
    var normalizedOptions = {
      debug: options.debug || false,
      id: options.id || (0, _uuid.v4)().replace(/\x2D/g, '')
    }; // False alarm: indexer is a constant of type Symbol.
    // eslint-disable-next-line security/detect-object-injection

    _this[SYMBOL_DEVICE_INFO_DEFERRED] = (0, _pDeferEs.default)(); // False alarm: indexer is a constant of type Symbol.
    // eslint-disable-next-line security/detect-object-injection

    _this[SYMBOL_EVENTS] = new _Exports.EventSource(); // False alarm: indexer is a constant of type Symbol.
    // eslint-disable-next-line security/detect-object-injection

    _this[SYMBOL_FORMAT_DEFERRED] = (0, _pDeferEs.default)(); // False alarm: indexer is a constant of type Symbol.
    // eslint-disable-next-line security/detect-object-injection

    _this[SYMBOL_OPTIONS] = normalizedOptions;
    return _this;
  }

  (0, _createClass2.default)(CustomAudioInputStream, [{
    key: "events",
    get:
    /** Gets the event source for listening to events. */
    // ESLint: This code will only works in browsers other than IE11. Only works in ES5 is okay.
    // @ts-ignore Accessors are only available when targeting ECMAScript 5 and higher.ts(1056)
    function get() {
      // False alarm: indexer is a constant of type Symbol.
      // eslint-disable-next-line security/detect-object-injection
      return this[SYMBOL_EVENTS];
    }
    /** Gets the format of the audio stream. */
    // Speech SDK quirks: `AudioStreamFormatImpl` is internal implementation while `AudioStreamFormat` is public.
    //                    It is weird to expose `AudioStreamFormatImpl` instead of `AudioStreamFormat`.
    // Speech SDK quirks: It is weird to return a `Promise` in a property.
    // Speech SDK quirks: In normal speech recognition, getter of "format" is called only after "attach".
    //                    But in Direct Line Speech, it is called before "attach".
    // ESLint: This code will only works in browsers other than IE11. Only works in ES5 is okay.
    // @ts-ignore Accessors are only available when targeting ECMAScript 5 and higher.ts(1056)

  }, {
    key: "format",
    get: function get() {
      this.debug('Getting "format".'); // False alarm: indexer is a constant of type Symbol.
      // eslint-disable-next-line security/detect-object-injection

      return this[SYMBOL_FORMAT_DEFERRED].promise;
    }
    /** Gets the ID of this audio stream. */

  }, {
    key: "id",
    value: function id() {
      // False alarm: indexer is a constant of type Symbol.
      // eslint-disable-next-line security/detect-object-injection
      return this[SYMBOL_OPTIONS].id;
    }
    /** Emits an event. */
    // Speech SDK quirks: In JavaScript, onXxx means "listen to event XXX".
    //                    Instead, in Speech SDK, it means "emit event XXX".

  }, {
    key: "onEvent",
    value: function onEvent(event) {
      // False alarm: indexer is a constant of type Symbol.
      // eslint-disable-next-line security/detect-object-injection
      this[SYMBOL_EVENTS].onEvent(event);

      _Exports.Events.instance.onEvent(event);
    }
    /** Emits an `AudioSourceInitializingEvent`. */

  }, {
    key: "emitInitializing",
    value: function emitInitializing() {
      this.debug('Emitting "AudioSourceInitializingEvent".');
      this.onEvent(new _Exports.AudioSourceInitializingEvent(this.id()));
    }
    /** Emits an `AudioSourceReadyEvent`. */

  }, {
    key: "emitReady",
    value: function emitReady() {
      this.debug('Emitting "AudioSourceReadyEvent".');
      this.onEvent(new _Exports.AudioSourceReadyEvent(this.id()));
    }
    /** Emits an `AudioSourceErrorEvent`. */
    // Speech SDK quirks: Since "turnOn" is never called and "turnOff" does not work in Direct Line Speech, the "source error" event is not emitted at all.
    //                    Instead, we only emit "node error" event.

  }, {
    key: "emitError",
    value: function emitError(error) {
      this.debug('Emitting "AudioSourceErrorEvent".', {
        error: error
      }); // Speech SDK quirks: "error" is a string, instead of object of type "Error".

      this.onEvent(new _Exports.AudioSourceErrorEvent(this.id(), error.message));
    }
    /** Emits an `AudioStreamNodeAttachingEvent`. */

  }, {
    key: "emitNodeAttaching",
    value: function emitNodeAttaching(audioNodeId) {
      this.debug("Emitting \"AudioStreamNodeAttachingEvent\" for node \"".concat(audioNodeId, "\"."));
      this.onEvent(new _Exports.AudioStreamNodeAttachingEvent(this.id(), audioNodeId));
    }
    /** Emits an `AudioStreamNodeAttachedEvent`. */

  }, {
    key: "emitNodeAttached",
    value: function emitNodeAttached(audioNodeId) {
      this.debug("Emitting \"AudioStreamNodeAttachedEvent\" for node \"".concat(audioNodeId, "\"."));
      this.onEvent(new _Exports.AudioStreamNodeAttachedEvent(this.id(), audioNodeId));
    }
    /** Emits an `AudioStreamNodeErrorEvent`. */

  }, {
    key: "emitNodeError",
    value: function emitNodeError(audioNodeId, error) {
      this.debug("Emitting \"AudioStreamNodeErrorEvent\" for node \"".concat(audioNodeId, "\"."), {
        error: error
      }); // Speech SDK quirks: "error" is a string, instead of object of type "Error".

      this.onEvent(new _Exports.AudioStreamNodeErrorEvent(this.id(), audioNodeId, error.message));
    }
    /** Emits an `AudioStreamNodeDetachedEvent`. */

  }, {
    key: "emitNodeDetached",
    value: function emitNodeDetached(audioNodeId) {
      this.debug('Emitting "AudioStreamNodeDetachedEvent".');
      this.onEvent(new _Exports.AudioStreamNodeDetachedEvent(this.id(), audioNodeId));
    }
    /** Emits an `AudioSourceOffEvent`. */

  }, {
    key: "emitOff",
    value: function emitOff() {
      this.debug('Emitting "AudioSourceOffEvent".');
      this.onEvent(new _Exports.AudioSourceOffEvent(this.id()));
    } // Speech SDK quirks: Although "close" is marked as abstract, it is never called in our observations.
    // ESLint: Speech SDK requires this function, but we are not implementing it.

  }, {
    key: "close",
    value: function close() {
      this.debug('Callback for "close".');
      throw new Error('Not implemented');
    } // Speech SDK quirks: Although "turnOn" is implemented in Speech SDK Push/PullAudioInputStream, it is never called in our observations.

  }, {
    key: "turnOn",
    value: function turnOn() {
      this.debug('Callback for "turnOn".');
      throw new Error('Not implemented');
    } // Speech SDK quirks: Although "detach" is implemented in Speech SDK Push/PullAudioInputStream, it is never called in our observations.

  }, {
    key: "detach",
    value: function detach() {
      this.debug('Callback for "detach".');
      throw new Error('Not implemented');
    }
    /** Log the message to console if `debug` is set to `true`. */

  }, {
    key: "debug",
    value: function debug(message) {
      var _console;

      for (var _len = arguments.length, args = new Array(_len > 1 ? _len - 1 : 0), _key = 1; _key < _len; _key++) {
        args[_key - 1] = arguments[_key];
      }

      // ESLint: For debugging, will only log when "debug" is set to "true".
      // False alarm: indexer is a constant of type Symbol.
      // eslint-disable-next-line no-console, security/detect-object-injection
      this[SYMBOL_OPTIONS].debug && (_console = console).info.apply(_console, ["CustomAudioInputStream: ".concat(message)].concat(args));
    }
    /** Implements this function. When called, it should start recording and return an `IAudioStreamNode`. */

  }, {
    key: "attach",
    value:
    /** Attaches the device by returning an audio node. */
    function attach(audioNodeId) {
      var _this2 = this;

      this.debug("Callback for \"attach\" with \"".concat(audioNodeId, "\"."));
      this.emitNodeAttaching(audioNodeId);
      return Promise.resolve().then( /*#__PURE__*/(0, _asyncToGenerator2.default)( /*#__PURE__*/_regenerator.default.mark(function _callee2() {
        var _yield$_this2$perform, audioStreamNode, deviceInfo, format;

        return _regenerator.default.wrap(function _callee2$(_context2) {
          while (1) {
            switch (_context2.prev = _context2.next) {
              case 0:
                _this2.emitInitializing();

                _context2.prev = 1;
                _context2.next = 4;
                return _this2.performAttach(audioNodeId);

              case 4:
                _yield$_this2$perform = _context2.sent;
                audioStreamNode = _yield$_this2$perform.audioStreamNode;
                deviceInfo = _yield$_this2$perform.deviceInfo;
                format = _yield$_this2$perform.format;

                // Although only getter of "format" is called before "attach" (in Direct Line Speech),
                // we are handling both "deviceInfo" and "format" in similar way for uniformity.
                // False alarm: indexer is a constant of type Symbol.
                // eslint-disable-next-line security/detect-object-injection
                _this2[SYMBOL_DEVICE_INFO_DEFERRED].resolve(deviceInfo); // False alarm: indexer is a constant of type Symbol.
                // eslint-disable-next-line security/detect-object-injection


                _this2[SYMBOL_FORMAT_DEFERRED].resolve(new _AudioStreamFormat.AudioStreamFormatImpl(format.samplesPerSec, format.bitsPerSample, format.channels));

                _this2.emitReady();

                _this2.emitNodeAttached(audioNodeId);

                return _context2.abrupt("return", {
                  detach: function () {
                    var _detach = (0, _asyncToGenerator2.default)( /*#__PURE__*/_regenerator.default.mark(function _callee() {
                      return _regenerator.default.wrap(function _callee$(_context) {
                        while (1) {
                          switch (_context.prev = _context.next) {
                            case 0:
                              _this2.debug("Detaching audio node \"".concat(audioNodeId, "\"."));

                              _context.next = 3;
                              return audioStreamNode.detach();

                            case 3:
                              // Speech SDK quirks: Since "turnOff" is not called in Direct Line Speech, we will emit event "source off" here instead.
                              _this2.emitOff();

                              _this2.emitNodeDetached(audioNodeId);

                            case 5:
                            case "end":
                              return _context.stop();
                          }
                        }
                      }, _callee);
                    }));

                    function detach() {
                      return _detach.apply(this, arguments);
                    }

                    return detach;
                  }(),
                  id: function id() {
                    return audioStreamNode.id();
                  },
                  read: function read() {
                    _this2.debug('Reading');

                    return audioStreamNode.read();
                  }
                });

              case 15:
                _context2.prev = 15;
                _context2.t0 = _context2["catch"](1);

                _this2.emitNodeError(audioNodeId, _context2.t0);

                throw _context2.t0;

              case 19:
              case "end":
                return _context2.stop();
            }
          }
        }, _callee2, null, [[1, 15]]);
      })));
    }
    /**
     * Implements this function. When called, it should stop recording. This is called before the `IAudioStreamNode.detach` function.
     *
     * Note: when using with Direct Line Speech, this function is never called.
     */
    // ESLint: We are not implementing this function because it is not called by Direct Line Speech.
    // eslint-disable-next-line class-methods-use-this

  }, {
    key: "performTurnOff",
    value: function performTurnOff() {
      // ESLint: "return" is required by TypeScript
      // eslint-disable-next-line no-useless-return
      return;
    }
    /** Turn off the audio device. This is called before detaching from the graph. */
    // Speech SDK quirks: It is confused to have both "turnOff" and "detach". "turnOff" is called before "detach".
    //                    Why don't we put all logics at "detach"?
    // Speech SDK quirks: Direct Line Speech never call "turnOff". "Source off" event need to be emitted during "detach" instead.
    //                    Also, custom implementation should be done at "detach" instead, such as ending and closing output streams.

  }, {
    key: "turnOff",
    value: function () {
      var _turnOff = (0, _asyncToGenerator2.default)( /*#__PURE__*/_regenerator.default.mark(function _callee3() {
        return _regenerator.default.wrap(function _callee3$(_context3) {
          while (1) {
            switch (_context3.prev = _context3.next) {
              case 0:
                this.debug("Callback for \"turnOff\".");
                _context3.next = 3;
                return this.performTurnOff();

              case 3:
              case "end":
                return _context3.stop();
            }
          }
        }, _callee3, this);
      }));

      function turnOff() {
        return _turnOff.apply(this, arguments);
      }

      return turnOff;
    }()
    /** Gets the device information. */
    // ESLint: This code will only works in browsers other than IE11. Only works in ES5 is okay.
    // @ts-ignore Accessors are only available when targeting ECMAScript 5 and higher.ts(1056)

  }, {
    key: "deviceInfo",
    get: function get() {
      this.debug("Getting \"deviceInfo\"."); // False alarm: indexer is a constant of type Symbol.
      // eslint-disable-next-line security/detect-object-injection

      return Promise.all([this[SYMBOL_DEVICE_INFO_DEFERRED].promise, this[SYMBOL_FORMAT_DEFERRED].promise]).then(function (_ref2) {
        var _ref3 = (0, _slicedToArray2.default)(_ref2, 2),
            _ref3$ = _ref3[0],
            connectivity = _ref3$.connectivity,
            manufacturer = _ref3$.manufacturer,
            model = _ref3$.model,
            type = _ref3$.type,
            _ref3$2 = _ref3[1],
            bitsPerSample = _ref3$2.bitsPerSample,
            channels = _ref3$2.channels,
            samplesPerSec = _ref3$2.samplesPerSec;

        return {
          bitspersample: bitsPerSample,
          channelcount: channels,
          connectivity: typeof connectivity === 'string' && !(0, _botframeworkWebchatCore.isForbiddenPropertyName)(connectivity) ? // Mitigated through denylisting.
          // eslint-disable-next-line security/detect-object-injection
          _Exports2.connectivity[connectivity] : connectivity || _Exports2.connectivity.Unknown,
          manufacturer: manufacturer || '',
          model: model || '',
          samplerate: samplesPerSec,
          // Mitigated through denylisting.
          // eslint-disable-next-line security/detect-object-injection
          type: typeof type === 'string' && !(0, _botframeworkWebchatCore.isForbiddenPropertyName)(type) ? _Exports2.type[type] : type || _Exports2.type.Unknown
        };
      });
    }
  }]);
  return CustomAudioInputStream;
}(_microsoftCognitiveservicesSpeechSdk.AudioInputStream);

var _default = CustomAudioInputStream;
exports.default = _default;
//# sourceMappingURL=data:application/json;charset=utf-8;base64,eyJ2ZXJzaW9uIjozLCJuYW1lcyI6WyJTWU1CT0xfREVWSUNFX0lORk9fREVGRVJSRUQiLCJTeW1ib2wiLCJTWU1CT0xfRVZFTlRTIiwiU1lNQk9MX0ZPUk1BVF9ERUZFUlJFRCIsIlNZTUJPTF9PUFRJT05TIiwiQ3VzdG9tQXVkaW9JbnB1dFN0cmVhbSIsIm9wdGlvbnMiLCJub3JtYWxpemVkT3B0aW9ucyIsImRlYnVnIiwiaWQiLCJ2NCIsInJlcGxhY2UiLCJjcmVhdGVEZWZlcnJlZCIsIkV2ZW50U291cmNlIiwicHJvbWlzZSIsImV2ZW50Iiwib25FdmVudCIsIkV2ZW50cyIsImluc3RhbmNlIiwiQXVkaW9Tb3VyY2VJbml0aWFsaXppbmdFdmVudCIsIkF1ZGlvU291cmNlUmVhZHlFdmVudCIsImVycm9yIiwiQXVkaW9Tb3VyY2VFcnJvckV2ZW50IiwibWVzc2FnZSIsImF1ZGlvTm9kZUlkIiwiQXVkaW9TdHJlYW1Ob2RlQXR0YWNoaW5nRXZlbnQiLCJBdWRpb1N0cmVhbU5vZGVBdHRhY2hlZEV2ZW50IiwiQXVkaW9TdHJlYW1Ob2RlRXJyb3JFdmVudCIsIkF1ZGlvU3RyZWFtTm9kZURldGFjaGVkRXZlbnQiLCJBdWRpb1NvdXJjZU9mZkV2ZW50IiwiRXJyb3IiLCJhcmdzIiwiY29uc29sZSIsImluZm8iLCJlbWl0Tm9kZUF0dGFjaGluZyIsIlByb21pc2UiLCJyZXNvbHZlIiwidGhlbiIsImVtaXRJbml0aWFsaXppbmciLCJwZXJmb3JtQXR0YWNoIiwiYXVkaW9TdHJlYW1Ob2RlIiwiZGV2aWNlSW5mbyIsImZvcm1hdCIsIkF1ZGlvU3RyZWFtRm9ybWF0SW1wbCIsInNhbXBsZXNQZXJTZWMiLCJiaXRzUGVyU2FtcGxlIiwiY2hhbm5lbHMiLCJlbWl0UmVhZHkiLCJlbWl0Tm9kZUF0dGFjaGVkIiwiZGV0YWNoIiwiZW1pdE9mZiIsImVtaXROb2RlRGV0YWNoZWQiLCJyZWFkIiwiZW1pdE5vZGVFcnJvciIsInBlcmZvcm1UdXJuT2ZmIiwiYWxsIiwiY29ubmVjdGl2aXR5IiwibWFudWZhY3R1cmVyIiwibW9kZWwiLCJ0eXBlIiwiYml0c3BlcnNhbXBsZSIsImNoYW5uZWxjb3VudCIsImlzRm9yYmlkZGVuUHJvcGVydHlOYW1lIiwiQ29ubmVjdGl2aXR5IiwiVW5rbm93biIsInNhbXBsZXJhdGUiLCJUeXBlIiwiQXVkaW9JbnB1dFN0cmVhbSJdLCJzb3VyY2VSb290IjoiYnVuZGxlOi8vLyIsInNvdXJjZXMiOlsiLi4vLi4vc3JjL3NwZWVjaC9DdXN0b21BdWRpb0lucHV0U3RyZWFtLnRzIl0sInNvdXJjZXNDb250ZW50IjpbImltcG9ydCB7IEF1ZGlvSW5wdXRTdHJlYW0gfSBmcm9tICdtaWNyb3NvZnQtY29nbml0aXZlc2VydmljZXMtc3BlZWNoLXNkayc7XG5cbmltcG9ydCB7XG4gIEF1ZGlvU291cmNlRXJyb3JFdmVudCxcbiAgQXVkaW9Tb3VyY2VFdmVudCxcbiAgQXVkaW9Tb3VyY2VJbml0aWFsaXppbmdFdmVudCxcbiAgQXVkaW9Tb3VyY2VPZmZFdmVudCxcbiAgQXVkaW9Tb3VyY2VSZWFkeUV2ZW50LFxuICBBdWRpb1N0cmVhbU5vZGVBdHRhY2hlZEV2ZW50LFxuICBBdWRpb1N0cmVhbU5vZGVBdHRhY2hpbmdFdmVudCxcbiAgQXVkaW9TdHJlYW1Ob2RlRGV0YWNoZWRFdmVudCxcbiAgQXVkaW9TdHJlYW1Ob2RlRXJyb3JFdmVudCxcbiAgRXZlbnRzLFxuICBFdmVudFNvdXJjZVxufSBmcm9tICdtaWNyb3NvZnQtY29nbml0aXZlc2VydmljZXMtc3BlZWNoLXNkay9kaXN0cmliL2xpYi9zcmMvY29tbW9uL0V4cG9ydHMnO1xuXG5pbXBvcnQgeyBBdWRpb1N0cmVhbUZvcm1hdEltcGwgfSBmcm9tICdtaWNyb3NvZnQtY29nbml0aXZlc2VydmljZXMtc3BlZWNoLXNkay9kaXN0cmliL2xpYi9zcmMvc2RrL0F1ZGlvL0F1ZGlvU3RyZWFtRm9ybWF0JztcblxuaW1wb3J0IHtcbiAgY29ubmVjdGl2aXR5IGFzIENvbm5lY3Rpdml0eSxcbiAgSVNwZWVjaENvbmZpZ0F1ZGlvRGV2aWNlLFxuICB0eXBlIGFzIFR5cGVcbn0gZnJvbSAnbWljcm9zb2Z0LWNvZ25pdGl2ZXNlcnZpY2VzLXNwZWVjaC1zZGsvZGlzdHJpYi9saWIvc3JjL2NvbW1vbi5zcGVlY2gvRXhwb3J0cyc7XG5cbmltcG9ydCB7IGlzRm9yYmlkZGVuUHJvcGVydHlOYW1lIH0gZnJvbSAnYm90ZnJhbWV3b3JrLXdlYmNoYXQtY29yZSc7XG5pbXBvcnQgeyB2NCB9IGZyb20gJ3V1aWQnO1xuaW1wb3J0IGNyZWF0ZURlZmVycmVkLCB7IERlZmVycmVkUHJvbWlzZSB9IGZyb20gJ3AtZGVmZXItZXM1JztcblxudHlwZSBBdWRpb1N0cmVhbU5vZGUgPSB7XG4gIGRldGFjaDogKCkgPT4gUHJvbWlzZTx2b2lkPjtcbiAgaWQ6ICgpID0+IHN0cmluZztcbiAgcmVhZDogKCkgPT4gUHJvbWlzZTxTdHJlYW1DaHVuazxBcnJheUJ1ZmZlcj4+O1xufTtcblxudHlwZSBEZXZpY2VJbmZvID0ge1xuICBjb25uZWN0aXZpdHk/OiBDb25uZWN0aXZpdHkgfCAnQmx1ZXRvb3RoJyB8ICdXaXJlZCcgfCAnV2lGaScgfCAnQ2VsbHVsYXInIHwgJ0luQnVpbHQnIHwgJ1Vua25vd24nO1xuICBtYW51ZmFjdHVyZXI/OiBzdHJpbmc7XG4gIG1vZGVsPzogc3RyaW5nO1xuICB0eXBlPzpcbiAgICB8IFR5cGVcbiAgICB8ICdQaG9uZSdcbiAgICB8ICdTcGVha2VyJ1xuICAgIHwgJ0NhcidcbiAgICB8ICdIZWFkc2V0J1xuICAgIHwgJ1RoZXJtb3N0YXQnXG4gICAgfCAnTWljcm9waG9uZXMnXG4gICAgfCAnRGVza3Bob25lJ1xuICAgIHwgJ1JlbW90ZUNvbnRyb2wnXG4gICAgfCAnVW5rbm93bidcbiAgICB8ICdGaWxlJ1xuICAgIHwgJ1N0cmVhbSc7XG59O1xuXG50eXBlIEZvcm1hdCA9IHtcbiAgYml0c1BlclNhbXBsZTogbnVtYmVyO1xuICBjaGFubmVsczogbnVtYmVyO1xuICBzYW1wbGVzUGVyU2VjOiBudW1iZXI7XG59O1xuXG50eXBlIE5vcm1hbGl6ZWRPcHRpb25zID0gUmVxdWlyZWQ8T21pdDxPcHRpb25zLCAnZGVidWcnPj4gJiB7XG4gIGRlYnVnOiBib29sZWFuO1xufTtcblxudHlwZSBPcHRpb25zID0ge1xuICBkZWJ1Zz86IHRydWU7XG4gIGlkPzogc3RyaW5nO1xufTtcblxudHlwZSBTdHJlYW1DaHVuazxUPiA9IHtcbiAgaXNFbmQ6IGJvb2xlYW47XG4gIGJ1ZmZlcjogVDtcbiAgdGltZVJlY2VpdmVkOiBudW1iZXI7XG59O1xuXG5jb25zdCBTWU1CT0xfREVWSUNFX0lORk9fREVGRVJSRUQgPSBTeW1ib2woJ2RldmljZUluZm9EZWZlcnJlZCcpO1xuY29uc3QgU1lNQk9MX0VWRU5UUyA9IFN5bWJvbCgnZXZlbnRzJyk7XG5jb25zdCBTWU1CT0xfRk9STUFUX0RFRkVSUkVEID0gU3ltYm9sKCdmb3JtYXREZWZlcnJlZCcpO1xuY29uc3QgU1lNQk9MX09QVElPTlMgPSBTeW1ib2woJ29wdGlvbnMnKTtcblxuLy8gU3BlZWNoIFNESyBxdWlya3M6IE9ubHkgMiBsaWZlY3ljbGUgZnVuY3Rpb25zIGFyZSBhY3R1YWxseSB1c2VkLlxuLy8gICAgICAgICAgICAgICAgICAgIFRoZXkgYXJlOiBhdHRhY2goKSBhbmQgdHVybk9mZigpLlxuLy8gICAgICAgICAgICAgICAgICAgIE90aGVycyBhcmUgbm90IHVzZWQsIGluY2x1ZGluZzogYmxvYigpLCBjbG9zZSgpLCBkZXRhY2goKSwgdHVybk9uKCkuXG5hYnN0cmFjdCBjbGFzcyBDdXN0b21BdWRpb0lucHV0U3RyZWFtIGV4dGVuZHMgQXVkaW9JbnB1dFN0cmVhbSB7XG4gIGNvbnN0cnVjdG9yKG9wdGlvbnM6IE9wdGlvbnMgPSB7fSkge1xuICAgIHN1cGVyKCk7XG5cbiAgICBjb25zdCBub3JtYWxpemVkT3B0aW9uczogTm9ybWFsaXplZE9wdGlvbnMgPSB7XG4gICAgICBkZWJ1Zzogb3B0aW9ucy5kZWJ1ZyB8fCBmYWxzZSxcbiAgICAgIGlkOiBvcHRpb25zLmlkIHx8IHY0KCkucmVwbGFjZSgvLS9ndSwgJycpXG4gICAgfTtcblxuICAgIC8vIEZhbHNlIGFsYXJtOiBpbmRleGVyIGlzIGEgY29uc3RhbnQgb2YgdHlwZSBTeW1ib2wuXG4gICAgLy8gZXNsaW50LWRpc2FibGUtbmV4dC1saW5lIHNlY3VyaXR5L2RldGVjdC1vYmplY3QtaW5qZWN0aW9uXG4gICAgdGhpc1tTWU1CT0xfREVWSUNFX0lORk9fREVGRVJSRURdID0gY3JlYXRlRGVmZXJyZWQ8RGV2aWNlSW5mbz4oKTtcblxuICAgIC8vIEZhbHNlIGFsYXJtOiBpbmRleGVyIGlzIGEgY29uc3RhbnQgb2YgdHlwZSBTeW1ib2wuXG4gICAgLy8gZXNsaW50LWRpc2FibGUtbmV4dC1saW5lIHNlY3VyaXR5L2RldGVjdC1vYmplY3QtaW5qZWN0aW9uXG4gICAgdGhpc1tTWU1CT0xfRVZFTlRTXSA9IG5ldyBFdmVudFNvdXJjZTxBdWRpb1NvdXJjZUV2ZW50PigpO1xuXG4gICAgLy8gRmFsc2UgYWxhcm06IGluZGV4ZXIgaXMgYSBjb25zdGFudCBvZiB0eXBlIFN5bWJvbC5cbiAgICAvLyBlc2xpbnQtZGlzYWJsZS1uZXh0LWxpbmUgc2VjdXJpdHkvZGV0ZWN0LW9iamVjdC1pbmplY3Rpb25cbiAgICB0aGlzW1NZTUJPTF9GT1JNQVRfREVGRVJSRURdID0gY3JlYXRlRGVmZXJyZWQ8QXVkaW9TdHJlYW1Gb3JtYXRJbXBsPigpO1xuXG4gICAgLy8gRmFsc2UgYWxhcm06IGluZGV4ZXIgaXMgYSBjb25zdGFudCBvZiB0eXBlIFN5bWJvbC5cbiAgICAvLyBlc2xpbnQtZGlzYWJsZS1uZXh0LWxpbmUgc2VjdXJpdHkvZGV0ZWN0LW9iamVjdC1pbmplY3Rpb25cbiAgICB0aGlzW1NZTUJPTF9PUFRJT05TXSA9IG5vcm1hbGl6ZWRPcHRpb25zO1xuICB9XG5cbiAgW1NZTUJPTF9ERVZJQ0VfSU5GT19ERUZFUlJFRF06IERlZmVycmVkUHJvbWlzZTxEZXZpY2VJbmZvPjtcbiAgW1NZTUJPTF9FVkVOVFNdOiBFdmVudFNvdXJjZTxBdWRpb1NvdXJjZUV2ZW50PjtcbiAgW1NZTUJPTF9GT1JNQVRfREVGRVJSRURdOiBEZWZlcnJlZFByb21pc2U8QXVkaW9TdHJlYW1Gb3JtYXRJbXBsPjtcbiAgW1NZTUJPTF9PUFRJT05TXTogTm9ybWFsaXplZE9wdGlvbnM7XG5cbiAgLyoqIEdldHMgdGhlIGV2ZW50IHNvdXJjZSBmb3IgbGlzdGVuaW5nIHRvIGV2ZW50cy4gKi9cbiAgLy8gRVNMaW50OiBUaGlzIGNvZGUgd2lsbCBvbmx5IHdvcmtzIGluIGJyb3dzZXJzIG90aGVyIHRoYW4gSUUxMS4gT25seSB3b3JrcyBpbiBFUzUgaXMgb2theS5cbiAgLy8gQHRzLWlnbm9yZSBBY2Nlc3NvcnMgYXJlIG9ubHkgYXZhaWxhYmxlIHdoZW4gdGFyZ2V0aW5nIEVDTUFTY3JpcHQgNSBhbmQgaGlnaGVyLnRzKDEwNTYpXG4gIGdldCBldmVudHMoKTogRXZlbnRTb3VyY2U8QXVkaW9Tb3VyY2VFdmVudD4ge1xuICAgIC8vIEZhbHNlIGFsYXJtOiBpbmRleGVyIGlzIGEgY29uc3RhbnQgb2YgdHlwZSBTeW1ib2wuXG4gICAgLy8gZXNsaW50LWRpc2FibGUtbmV4dC1saW5lIHNlY3VyaXR5L2RldGVjdC1vYmplY3QtaW5qZWN0aW9uXG4gICAgcmV0dXJuIHRoaXNbU1lNQk9MX0VWRU5UU107XG4gIH1cblxuICAvKiogR2V0cyB0aGUgZm9ybWF0IG9mIHRoZSBhdWRpbyBzdHJlYW0uICovXG4gIC8vIFNwZWVjaCBTREsgcXVpcmtzOiBgQXVkaW9TdHJlYW1Gb3JtYXRJbXBsYCBpcyBpbnRlcm5hbCBpbXBsZW1lbnRhdGlvbiB3aGlsZSBgQXVkaW9TdHJlYW1Gb3JtYXRgIGlzIHB1YmxpYy5cbiAgLy8gICAgICAgICAgICAgICAgICAgIEl0IGlzIHdlaXJkIHRvIGV4cG9zZSBgQXVkaW9TdHJlYW1Gb3JtYXRJbXBsYCBpbnN0ZWFkIG9mIGBBdWRpb1N0cmVhbUZvcm1hdGAuXG4gIC8vIFNwZWVjaCBTREsgcXVpcmtzOiBJdCBpcyB3ZWlyZCB0byByZXR1cm4gYSBgUHJvbWlzZWAgaW4gYSBwcm9wZXJ0eS5cbiAgLy8gU3BlZWNoIFNESyBxdWlya3M6IEluIG5vcm1hbCBzcGVlY2ggcmVjb2duaXRpb24sIGdldHRlciBvZiBcImZvcm1hdFwiIGlzIGNhbGxlZCBvbmx5IGFmdGVyIFwiYXR0YWNoXCIuXG4gIC8vICAgICAgICAgICAgICAgICAgICBCdXQgaW4gRGlyZWN0IExpbmUgU3BlZWNoLCBpdCBpcyBjYWxsZWQgYmVmb3JlIFwiYXR0YWNoXCIuXG4gIC8vIEVTTGludDogVGhpcyBjb2RlIHdpbGwgb25seSB3b3JrcyBpbiBicm93c2VycyBvdGhlciB0aGFuIElFMTEuIE9ubHkgd29ya3MgaW4gRVM1IGlzIG9rYXkuXG4gIC8vIEB0cy1pZ25vcmUgQWNjZXNzb3JzIGFyZSBvbmx5IGF2YWlsYWJsZSB3aGVuIHRhcmdldGluZyBFQ01BU2NyaXB0IDUgYW5kIGhpZ2hlci50cygxMDU2KVxuICBnZXQgZm9ybWF0KCk6IFByb21pc2U8QXVkaW9TdHJlYW1Gb3JtYXRJbXBsPiB7XG4gICAgdGhpcy5kZWJ1ZygnR2V0dGluZyBcImZvcm1hdFwiLicpO1xuXG4gICAgLy8gRmFsc2UgYWxhcm06IGluZGV4ZXIgaXMgYSBjb25zdGFudCBvZiB0eXBlIFN5bWJvbC5cbiAgICAvLyBlc2xpbnQtZGlzYWJsZS1uZXh0LWxpbmUgc2VjdXJpdHkvZGV0ZWN0LW9iamVjdC1pbmplY3Rpb25cbiAgICByZXR1cm4gdGhpc1tTWU1CT0xfRk9STUFUX0RFRkVSUkVEXS5wcm9taXNlO1xuICB9XG5cbiAgLyoqIEdldHMgdGhlIElEIG9mIHRoaXMgYXVkaW8gc3RyZWFtLiAqL1xuICBpZCgpOiBzdHJpbmcge1xuICAgIC8vIEZhbHNlIGFsYXJtOiBpbmRleGVyIGlzIGEgY29uc3RhbnQgb2YgdHlwZSBTeW1ib2wuXG4gICAgLy8gZXNsaW50LWRpc2FibGUtbmV4dC1saW5lIHNlY3VyaXR5L2RldGVjdC1vYmplY3QtaW5qZWN0aW9uXG4gICAgcmV0dXJuIHRoaXNbU1lNQk9MX09QVElPTlNdLmlkO1xuICB9XG5cbiAgLyoqIEVtaXRzIGFuIGV2ZW50LiAqL1xuICAvLyBTcGVlY2ggU0RLIHF1aXJrczogSW4gSmF2YVNjcmlwdCwgb25YeHggbWVhbnMgXCJsaXN0ZW4gdG8gZXZlbnQgWFhYXCIuXG4gIC8vICAgICAgICAgICAgICAgICAgICBJbnN0ZWFkLCBpbiBTcGVlY2ggU0RLLCBpdCBtZWFucyBcImVtaXQgZXZlbnQgWFhYXCIuXG4gIHByb3RlY3RlZCBvbkV2ZW50KGV2ZW50OiBBdWRpb1NvdXJjZUV2ZW50KTogdm9pZCB7XG4gICAgLy8gRmFsc2UgYWxhcm06IGluZGV4ZXIgaXMgYSBjb25zdGFudCBvZiB0eXBlIFN5bWJvbC5cbiAgICAvLyBlc2xpbnQtZGlzYWJsZS1uZXh0LWxpbmUgc2VjdXJpdHkvZGV0ZWN0LW9iamVjdC1pbmplY3Rpb25cbiAgICB0aGlzW1NZTUJPTF9FVkVOVFNdLm9uRXZlbnQoZXZlbnQpO1xuICAgIEV2ZW50cy5pbnN0YW5jZS5vbkV2ZW50KGV2ZW50KTtcbiAgfVxuXG4gIC8qKiBFbWl0cyBhbiBgQXVkaW9Tb3VyY2VJbml0aWFsaXppbmdFdmVudGAuICovXG4gIHByb3RlY3RlZCBlbWl0SW5pdGlhbGl6aW5nKCk6IHZvaWQge1xuICAgIHRoaXMuZGVidWcoJ0VtaXR0aW5nIFwiQXVkaW9Tb3VyY2VJbml0aWFsaXppbmdFdmVudFwiLicpO1xuICAgIHRoaXMub25FdmVudChuZXcgQXVkaW9Tb3VyY2VJbml0aWFsaXppbmdFdmVudCh0aGlzLmlkKCkpKTtcbiAgfVxuXG4gIC8qKiBFbWl0cyBhbiBgQXVkaW9Tb3VyY2VSZWFkeUV2ZW50YC4gKi9cbiAgcHJvdGVjdGVkIGVtaXRSZWFkeSgpOiB2b2lkIHtcbiAgICB0aGlzLmRlYnVnKCdFbWl0dGluZyBcIkF1ZGlvU291cmNlUmVhZHlFdmVudFwiLicpO1xuICAgIHRoaXMub25FdmVudChuZXcgQXVkaW9Tb3VyY2VSZWFkeUV2ZW50KHRoaXMuaWQoKSkpO1xuICB9XG5cbiAgLyoqIEVtaXRzIGFuIGBBdWRpb1NvdXJjZUVycm9yRXZlbnRgLiAqL1xuICAvLyBTcGVlY2ggU0RLIHF1aXJrczogU2luY2UgXCJ0dXJuT25cIiBpcyBuZXZlciBjYWxsZWQgYW5kIFwidHVybk9mZlwiIGRvZXMgbm90IHdvcmsgaW4gRGlyZWN0IExpbmUgU3BlZWNoLCB0aGUgXCJzb3VyY2UgZXJyb3JcIiBldmVudCBpcyBub3QgZW1pdHRlZCBhdCBhbGwuXG4gIC8vICAgICAgICAgICAgICAgICAgICBJbnN0ZWFkLCB3ZSBvbmx5IGVtaXQgXCJub2RlIGVycm9yXCIgZXZlbnQuXG4gIHByb3RlY3RlZCBlbWl0RXJyb3IoZXJyb3I6IEVycm9yKTogdm9pZCB7XG4gICAgdGhpcy5kZWJ1ZygnRW1pdHRpbmcgXCJBdWRpb1NvdXJjZUVycm9yRXZlbnRcIi4nLCB7IGVycm9yIH0pO1xuXG4gICAgLy8gU3BlZWNoIFNESyBxdWlya3M6IFwiZXJyb3JcIiBpcyBhIHN0cmluZywgaW5zdGVhZCBvZiBvYmplY3Qgb2YgdHlwZSBcIkVycm9yXCIuXG4gICAgdGhpcy5vbkV2ZW50KG5ldyBBdWRpb1NvdXJjZUVycm9yRXZlbnQodGhpcy5pZCgpLCBlcnJvci5tZXNzYWdlKSk7XG4gIH1cblxuICAvKiogRW1pdHMgYW4gYEF1ZGlvU3RyZWFtTm9kZUF0dGFjaGluZ0V2ZW50YC4gKi9cbiAgcHJvdGVjdGVkIGVtaXROb2RlQXR0YWNoaW5nKGF1ZGlvTm9kZUlkOiBzdHJpbmcpOiB2b2lkIHtcbiAgICB0aGlzLmRlYnVnKGBFbWl0dGluZyBcIkF1ZGlvU3RyZWFtTm9kZUF0dGFjaGluZ0V2ZW50XCIgZm9yIG5vZGUgXCIke2F1ZGlvTm9kZUlkfVwiLmApO1xuICAgIHRoaXMub25FdmVudChuZXcgQXVkaW9TdHJlYW1Ob2RlQXR0YWNoaW5nRXZlbnQodGhpcy5pZCgpLCBhdWRpb05vZGVJZCkpO1xuICB9XG5cbiAgLyoqIEVtaXRzIGFuIGBBdWRpb1N0cmVhbU5vZGVBdHRhY2hlZEV2ZW50YC4gKi9cbiAgcHJvdGVjdGVkIGVtaXROb2RlQXR0YWNoZWQoYXVkaW9Ob2RlSWQ6IHN0cmluZyk6IHZvaWQge1xuICAgIHRoaXMuZGVidWcoYEVtaXR0aW5nIFwiQXVkaW9TdHJlYW1Ob2RlQXR0YWNoZWRFdmVudFwiIGZvciBub2RlIFwiJHthdWRpb05vZGVJZH1cIi5gKTtcbiAgICB0aGlzLm9uRXZlbnQobmV3IEF1ZGlvU3RyZWFtTm9kZUF0dGFjaGVkRXZlbnQodGhpcy5pZCgpLCBhdWRpb05vZGVJZCkpO1xuICB9XG5cbiAgLyoqIEVtaXRzIGFuIGBBdWRpb1N0cmVhbU5vZGVFcnJvckV2ZW50YC4gKi9cbiAgcHJvdGVjdGVkIGVtaXROb2RlRXJyb3IoYXVkaW9Ob2RlSWQ6IHN0cmluZywgZXJyb3I6IEVycm9yKTogdm9pZCB7XG4gICAgdGhpcy5kZWJ1ZyhgRW1pdHRpbmcgXCJBdWRpb1N0cmVhbU5vZGVFcnJvckV2ZW50XCIgZm9yIG5vZGUgXCIke2F1ZGlvTm9kZUlkfVwiLmAsIHsgZXJyb3IgfSk7XG5cbiAgICAvLyBTcGVlY2ggU0RLIHF1aXJrczogXCJlcnJvclwiIGlzIGEgc3RyaW5nLCBpbnN0ZWFkIG9mIG9iamVjdCBvZiB0eXBlIFwiRXJyb3JcIi5cbiAgICB0aGlzLm9uRXZlbnQobmV3IEF1ZGlvU3RyZWFtTm9kZUVycm9yRXZlbnQodGhpcy5pZCgpLCBhdWRpb05vZGVJZCwgZXJyb3IubWVzc2FnZSkpO1xuICB9XG5cbiAgLyoqIEVtaXRzIGFuIGBBdWRpb1N0cmVhbU5vZGVEZXRhY2hlZEV2ZW50YC4gKi9cbiAgcHJvdGVjdGVkIGVtaXROb2RlRGV0YWNoZWQoYXVkaW9Ob2RlSWQ6IHN0cmluZyk6IHZvaWQge1xuICAgIHRoaXMuZGVidWcoJ0VtaXR0aW5nIFwiQXVkaW9TdHJlYW1Ob2RlRGV0YWNoZWRFdmVudFwiLicpO1xuICAgIHRoaXMub25FdmVudChuZXcgQXVkaW9TdHJlYW1Ob2RlRGV0YWNoZWRFdmVudCh0aGlzLmlkKCksIGF1ZGlvTm9kZUlkKSk7XG4gIH1cblxuICAvKiogRW1pdHMgYW4gYEF1ZGlvU291cmNlT2ZmRXZlbnRgLiAqL1xuICBwcm90ZWN0ZWQgZW1pdE9mZigpOiB2b2lkIHtcbiAgICB0aGlzLmRlYnVnKCdFbWl0dGluZyBcIkF1ZGlvU291cmNlT2ZmRXZlbnRcIi4nKTtcbiAgICB0aGlzLm9uRXZlbnQobmV3IEF1ZGlvU291cmNlT2ZmRXZlbnQodGhpcy5pZCgpKSk7XG4gIH1cblxuICAvLyBTcGVlY2ggU0RLIHF1aXJrczogQWx0aG91Z2ggXCJjbG9zZVwiIGlzIG1hcmtlZCBhcyBhYnN0cmFjdCwgaXQgaXMgbmV2ZXIgY2FsbGVkIGluIG91ciBvYnNlcnZhdGlvbnMuXG4gIC8vIEVTTGludDogU3BlZWNoIFNESyByZXF1aXJlcyB0aGlzIGZ1bmN0aW9uLCBidXQgd2UgYXJlIG5vdCBpbXBsZW1lbnRpbmcgaXQuXG4gIGNsb3NlKCk6IHZvaWQge1xuICAgIHRoaXMuZGVidWcoJ0NhbGxiYWNrIGZvciBcImNsb3NlXCIuJyk7XG5cbiAgICB0aHJvdyBuZXcgRXJyb3IoJ05vdCBpbXBsZW1lbnRlZCcpO1xuICB9XG5cbiAgLy8gU3BlZWNoIFNESyBxdWlya3M6IEFsdGhvdWdoIFwidHVybk9uXCIgaXMgaW1wbGVtZW50ZWQgaW4gU3BlZWNoIFNESyBQdXNoL1B1bGxBdWRpb0lucHV0U3RyZWFtLCBpdCBpcyBuZXZlciBjYWxsZWQgaW4gb3VyIG9ic2VydmF0aW9ucy5cbiAgdHVybk9uKCk6IHZvaWQge1xuICAgIHRoaXMuZGVidWcoJ0NhbGxiYWNrIGZvciBcInR1cm5PblwiLicpO1xuXG4gICAgdGhyb3cgbmV3IEVycm9yKCdOb3QgaW1wbGVtZW50ZWQnKTtcbiAgfVxuXG4gIC8vIFNwZWVjaCBTREsgcXVpcmtzOiBBbHRob3VnaCBcImRldGFjaFwiIGlzIGltcGxlbWVudGVkIGluIFNwZWVjaCBTREsgUHVzaC9QdWxsQXVkaW9JbnB1dFN0cmVhbSwgaXQgaXMgbmV2ZXIgY2FsbGVkIGluIG91ciBvYnNlcnZhdGlvbnMuXG4gIGRldGFjaCgpOiB2b2lkIHtcbiAgICB0aGlzLmRlYnVnKCdDYWxsYmFjayBmb3IgXCJkZXRhY2hcIi4nKTtcblxuICAgIHRocm93IG5ldyBFcnJvcignTm90IGltcGxlbWVudGVkJyk7XG4gIH1cblxuICAvKiogTG9nIHRoZSBtZXNzYWdlIHRvIGNvbnNvbGUgaWYgYGRlYnVnYCBpcyBzZXQgdG8gYHRydWVgLiAqL1xuICBwcml2YXRlIGRlYnVnKG1lc3NhZ2UsIC4uLmFyZ3MpIHtcbiAgICAvLyBFU0xpbnQ6IEZvciBkZWJ1Z2dpbmcsIHdpbGwgb25seSBsb2cgd2hlbiBcImRlYnVnXCIgaXMgc2V0IHRvIFwidHJ1ZVwiLlxuICAgIC8vIEZhbHNlIGFsYXJtOiBpbmRleGVyIGlzIGEgY29uc3RhbnQgb2YgdHlwZSBTeW1ib2wuXG4gICAgLy8gZXNsaW50LWRpc2FibGUtbmV4dC1saW5lIG5vLWNvbnNvbGUsIHNlY3VyaXR5L2RldGVjdC1vYmplY3QtaW5qZWN0aW9uXG4gICAgdGhpc1tTWU1CT0xfT1BUSU9OU10uZGVidWcgJiYgY29uc29sZS5pbmZvKGBDdXN0b21BdWRpb0lucHV0U3RyZWFtOiAke21lc3NhZ2V9YCwgLi4uYXJncyk7XG4gIH1cblxuICAvKiogSW1wbGVtZW50cyB0aGlzIGZ1bmN0aW9uLiBXaGVuIGNhbGxlZCwgaXQgc2hvdWxkIHN0YXJ0IHJlY29yZGluZyBhbmQgcmV0dXJuIGFuIGBJQXVkaW9TdHJlYW1Ob2RlYC4gKi9cbiAgcHJvdGVjdGVkIGFic3RyYWN0IHBlcmZvcm1BdHRhY2goYXVkaW9Ob2RlSWQ6IHN0cmluZyk6IFByb21pc2U8e1xuICAgIGF1ZGlvU3RyZWFtTm9kZTogQXVkaW9TdHJlYW1Ob2RlO1xuICAgIGRldmljZUluZm86IERldmljZUluZm87XG4gICAgZm9ybWF0OiBGb3JtYXQ7XG4gIH0+O1xuXG4gIC8qKiBBdHRhY2hlcyB0aGUgZGV2aWNlIGJ5IHJldHVybmluZyBhbiBhdWRpbyBub2RlLiAqL1xuICBhdHRhY2goYXVkaW9Ob2RlSWQ6IHN0cmluZyk6IFByb21pc2U8QXVkaW9TdHJlYW1Ob2RlPiB7XG4gICAgdGhpcy5kZWJ1ZyhgQ2FsbGJhY2sgZm9yIFwiYXR0YWNoXCIgd2l0aCBcIiR7YXVkaW9Ob2RlSWR9XCIuYCk7XG5cbiAgICB0aGlzLmVtaXROb2RlQXR0YWNoaW5nKGF1ZGlvTm9kZUlkKTtcblxuICAgIHJldHVybiBQcm9taXNlLnJlc29sdmUoKS50aGVuPEF1ZGlvU3RyZWFtTm9kZT4oYXN5bmMgKCkgPT4ge1xuICAgICAgdGhpcy5lbWl0SW5pdGlhbGl6aW5nKCk7XG5cbiAgICAgIHRyeSB7XG4gICAgICAgIGNvbnN0IHsgYXVkaW9TdHJlYW1Ob2RlLCBkZXZpY2VJbmZvLCBmb3JtYXQgfSA9IGF3YWl0IHRoaXMucGVyZm9ybUF0dGFjaChhdWRpb05vZGVJZCk7XG5cbiAgICAgICAgLy8gQWx0aG91Z2ggb25seSBnZXR0ZXIgb2YgXCJmb3JtYXRcIiBpcyBjYWxsZWQgYmVmb3JlIFwiYXR0YWNoXCIgKGluIERpcmVjdCBMaW5lIFNwZWVjaCksXG4gICAgICAgIC8vIHdlIGFyZSBoYW5kbGluZyBib3RoIFwiZGV2aWNlSW5mb1wiIGFuZCBcImZvcm1hdFwiIGluIHNpbWlsYXIgd2F5IGZvciB1bmlmb3JtaXR5LlxuXG4gICAgICAgIC8vIEZhbHNlIGFsYXJtOiBpbmRleGVyIGlzIGEgY29uc3RhbnQgb2YgdHlwZSBTeW1ib2wuXG4gICAgICAgIC8vIGVzbGludC1kaXNhYmxlLW5leHQtbGluZSBzZWN1cml0eS9kZXRlY3Qtb2JqZWN0LWluamVjdGlvblxuICAgICAgICB0aGlzW1NZTUJPTF9ERVZJQ0VfSU5GT19ERUZFUlJFRF0ucmVzb2x2ZShkZXZpY2VJbmZvKTtcblxuICAgICAgICAvLyBGYWxzZSBhbGFybTogaW5kZXhlciBpcyBhIGNvbnN0YW50IG9mIHR5cGUgU3ltYm9sLlxuICAgICAgICAvLyBlc2xpbnQtZGlzYWJsZS1uZXh0LWxpbmUgc2VjdXJpdHkvZGV0ZWN0LW9iamVjdC1pbmplY3Rpb25cbiAgICAgICAgdGhpc1tTWU1CT0xfRk9STUFUX0RFRkVSUkVEXS5yZXNvbHZlKFxuICAgICAgICAgIG5ldyBBdWRpb1N0cmVhbUZvcm1hdEltcGwoZm9ybWF0LnNhbXBsZXNQZXJTZWMsIGZvcm1hdC5iaXRzUGVyU2FtcGxlLCBmb3JtYXQuY2hhbm5lbHMpXG4gICAgICAgICk7XG5cbiAgICAgICAgdGhpcy5lbWl0UmVhZHkoKTtcbiAgICAgICAgdGhpcy5lbWl0Tm9kZUF0dGFjaGVkKGF1ZGlvTm9kZUlkKTtcblxuICAgICAgICByZXR1cm4ge1xuICAgICAgICAgIGRldGFjaDogYXN5bmMgKCkgPT4ge1xuICAgICAgICAgICAgdGhpcy5kZWJ1ZyhgRGV0YWNoaW5nIGF1ZGlvIG5vZGUgXCIke2F1ZGlvTm9kZUlkfVwiLmApO1xuXG4gICAgICAgICAgICBhd2FpdCBhdWRpb1N0cmVhbU5vZGUuZGV0YWNoKCk7XG5cbiAgICAgICAgICAgIC8vIFNwZWVjaCBTREsgcXVpcmtzOiBTaW5jZSBcInR1cm5PZmZcIiBpcyBub3QgY2FsbGVkIGluIERpcmVjdCBMaW5lIFNwZWVjaCwgd2Ugd2lsbCBlbWl0IGV2ZW50IFwic291cmNlIG9mZlwiIGhlcmUgaW5zdGVhZC5cbiAgICAgICAgICAgIHRoaXMuZW1pdE9mZigpO1xuICAgICAgICAgICAgdGhpcy5lbWl0Tm9kZURldGFjaGVkKGF1ZGlvTm9kZUlkKTtcbiAgICAgICAgICB9LFxuICAgICAgICAgIGlkOiAoKSA9PiBhdWRpb1N0cmVhbU5vZGUuaWQoKSxcbiAgICAgICAgICByZWFkOiAoKSA9PiB7XG4gICAgICAgICAgICB0aGlzLmRlYnVnKCdSZWFkaW5nJyk7XG5cbiAgICAgICAgICAgIHJldHVybiBhdWRpb1N0cmVhbU5vZGUucmVhZCgpO1xuICAgICAgICAgIH1cbiAgICAgICAgfTtcbiAgICAgIH0gY2F0Y2ggKGVycm9yKSB7XG4gICAgICAgIHRoaXMuZW1pdE5vZGVFcnJvcihhdWRpb05vZGVJZCwgZXJyb3IpO1xuXG4gICAgICAgIHRocm93IGVycm9yO1xuICAgICAgfVxuICAgIH0pO1xuICB9XG5cbiAgLyoqXG4gICAqIEltcGxlbWVudHMgdGhpcyBmdW5jdGlvbi4gV2hlbiBjYWxsZWQsIGl0IHNob3VsZCBzdG9wIHJlY29yZGluZy4gVGhpcyBpcyBjYWxsZWQgYmVmb3JlIHRoZSBgSUF1ZGlvU3RyZWFtTm9kZS5kZXRhY2hgIGZ1bmN0aW9uLlxuICAgKlxuICAgKiBOb3RlOiB3aGVuIHVzaW5nIHdpdGggRGlyZWN0IExpbmUgU3BlZWNoLCB0aGlzIGZ1bmN0aW9uIGlzIG5ldmVyIGNhbGxlZC5cbiAgICovXG5cbiAgLy8gRVNMaW50OiBXZSBhcmUgbm90IGltcGxlbWVudGluZyB0aGlzIGZ1bmN0aW9uIGJlY2F1c2UgaXQgaXMgbm90IGNhbGxlZCBieSBEaXJlY3QgTGluZSBTcGVlY2guXG4gIC8vIGVzbGludC1kaXNhYmxlLW5leHQtbGluZSBjbGFzcy1tZXRob2RzLXVzZS10aGlzXG4gIHByb3RlY3RlZCBwZXJmb3JtVHVybk9mZigpOiBQcm9taXNlPHZvaWQ+IHtcbiAgICAvLyBFU0xpbnQ6IFwicmV0dXJuXCIgaXMgcmVxdWlyZWQgYnkgVHlwZVNjcmlwdFxuICAgIC8vIGVzbGludC1kaXNhYmxlLW5leHQtbGluZSBuby11c2VsZXNzLXJldHVyblxuICAgIHJldHVybjtcbiAgfVxuXG4gIC8qKiBUdXJuIG9mZiB0aGUgYXVkaW8gZGV2aWNlLiBUaGlzIGlzIGNhbGxlZCBiZWZvcmUgZGV0YWNoaW5nIGZyb20gdGhlIGdyYXBoLiAqL1xuICAvLyBTcGVlY2ggU0RLIHF1aXJrczogSXQgaXMgY29uZnVzZWQgdG8gaGF2ZSBib3RoIFwidHVybk9mZlwiIGFuZCBcImRldGFjaFwiLiBcInR1cm5PZmZcIiBpcyBjYWxsZWQgYmVmb3JlIFwiZGV0YWNoXCIuXG4gIC8vICAgICAgICAgICAgICAgICAgICBXaHkgZG9uJ3Qgd2UgcHV0IGFsbCBsb2dpY3MgYXQgXCJkZXRhY2hcIj9cbiAgLy8gU3BlZWNoIFNESyBxdWlya3M6IERpcmVjdCBMaW5lIFNwZWVjaCBuZXZlciBjYWxsIFwidHVybk9mZlwiLiBcIlNvdXJjZSBvZmZcIiBldmVudCBuZWVkIHRvIGJlIGVtaXR0ZWQgZHVyaW5nIFwiZGV0YWNoXCIgaW5zdGVhZC5cbiAgLy8gICAgICAgICAgICAgICAgICAgIEFsc28sIGN1c3RvbSBpbXBsZW1lbnRhdGlvbiBzaG91bGQgYmUgZG9uZSBhdCBcImRldGFjaFwiIGluc3RlYWQsIHN1Y2ggYXMgZW5kaW5nIGFuZCBjbG9zaW5nIG91dHB1dCBzdHJlYW1zLlxuICBhc3luYyB0dXJuT2ZmKCk6IFByb21pc2U8dm9pZD4ge1xuICAgIHRoaXMuZGVidWcoYENhbGxiYWNrIGZvciBcInR1cm5PZmZcIi5gKTtcblxuICAgIGF3YWl0IHRoaXMucGVyZm9ybVR1cm5PZmYoKTtcbiAgfVxuXG4gIC8qKiBHZXRzIHRoZSBkZXZpY2UgaW5mb3JtYXRpb24uICovXG4gIC8vIEVTTGludDogVGhpcyBjb2RlIHdpbGwgb25seSB3b3JrcyBpbiBicm93c2VycyBvdGhlciB0aGFuIElFMTEuIE9ubHkgd29ya3MgaW4gRVM1IGlzIG9rYXkuXG4gIC8vIEB0cy1pZ25vcmUgQWNjZXNzb3JzIGFyZSBvbmx5IGF2YWlsYWJsZSB3aGVuIHRhcmdldGluZyBFQ01BU2NyaXB0IDUgYW5kIGhpZ2hlci50cygxMDU2KVxuICBnZXQgZGV2aWNlSW5mbygpOiBQcm9taXNlPElTcGVlY2hDb25maWdBdWRpb0RldmljZT4ge1xuICAgIHRoaXMuZGVidWcoYEdldHRpbmcgXCJkZXZpY2VJbmZvXCIuYCk7XG5cbiAgICAvLyBGYWxzZSBhbGFybTogaW5kZXhlciBpcyBhIGNvbnN0YW50IG9mIHR5cGUgU3ltYm9sLlxuICAgIC8vIGVzbGludC1kaXNhYmxlLW5leHQtbGluZSBzZWN1cml0eS9kZXRlY3Qtb2JqZWN0LWluamVjdGlvblxuICAgIHJldHVybiBQcm9taXNlLmFsbChbdGhpc1tTWU1CT0xfREVWSUNFX0lORk9fREVGRVJSRURdLnByb21pc2UsIHRoaXNbU1lNQk9MX0ZPUk1BVF9ERUZFUlJFRF0ucHJvbWlzZV0pLnRoZW4oXG4gICAgICAoW3sgY29ubmVjdGl2aXR5LCBtYW51ZmFjdHVyZXIsIG1vZGVsLCB0eXBlIH0sIHsgYml0c1BlclNhbXBsZSwgY2hhbm5lbHMsIHNhbXBsZXNQZXJTZWMgfV0pID0+ICh7XG4gICAgICAgIGJpdHNwZXJzYW1wbGU6IGJpdHNQZXJTYW1wbGUsXG4gICAgICAgIGNoYW5uZWxjb3VudDogY2hhbm5lbHMsXG4gICAgICAgIGNvbm5lY3Rpdml0eTpcbiAgICAgICAgICB0eXBlb2YgY29ubmVjdGl2aXR5ID09PSAnc3RyaW5nJyAmJiAhaXNGb3JiaWRkZW5Qcm9wZXJ0eU5hbWUoY29ubmVjdGl2aXR5KVxuICAgICAgICAgICAgPyAvLyBNaXRpZ2F0ZWQgdGhyb3VnaCBkZW55bGlzdGluZy5cbiAgICAgICAgICAgICAgLy8gZXNsaW50LWRpc2FibGUtbmV4dC1saW5lIHNlY3VyaXR5L2RldGVjdC1vYmplY3QtaW5qZWN0aW9uXG4gICAgICAgICAgICAgIENvbm5lY3Rpdml0eVtjb25uZWN0aXZpdHldXG4gICAgICAgICAgICA6IGNvbm5lY3Rpdml0eSB8fCBDb25uZWN0aXZpdHkuVW5rbm93bixcbiAgICAgICAgbWFudWZhY3R1cmVyOiBtYW51ZmFjdHVyZXIgfHwgJycsXG4gICAgICAgIG1vZGVsOiBtb2RlbCB8fCAnJyxcbiAgICAgICAgc2FtcGxlcmF0ZTogc2FtcGxlc1BlclNlYyxcbiAgICAgICAgLy8gTWl0aWdhdGVkIHRocm91Z2ggZGVueWxpc3RpbmcuXG4gICAgICAgIC8vIGVzbGludC1kaXNhYmxlLW5leHQtbGluZSBzZWN1cml0eS9kZXRlY3Qtb2JqZWN0LWluamVjdGlvblxuICAgICAgICB0eXBlOiB0eXBlb2YgdHlwZSA9PT0gJ3N0cmluZycgJiYgIWlzRm9yYmlkZGVuUHJvcGVydHlOYW1lKHR5cGUpID8gVHlwZVt0eXBlXSA6IHR5cGUgfHwgVHlwZS5Vbmtub3duXG4gICAgICB9KVxuICAgICk7XG4gIH1cbn1cblxuZXhwb3J0IGRlZmF1bHQgQ3VzdG9tQXVkaW9JbnB1dFN0cmVhbTtcblxuZXhwb3J0IHR5cGUgeyBBdWRpb1N0cmVhbU5vZGUsIERldmljZUluZm8sIEZvcm1hdCwgT3B0aW9ucyB9O1xuIl0sIm1hcHBpbmdzIjoiOzs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7OztBQUFBOztBQUVBOztBQWNBOztBQUVBOztBQU1BOztBQUNBOztBQUNBOzs7Ozs7QUFnREEsSUFBTUEsMkJBQTJCLEdBQUdDLE1BQU0sQ0FBQyxvQkFBRCxDQUExQztBQUNBLElBQU1DLGFBQWEsR0FBR0QsTUFBTSxDQUFDLFFBQUQsQ0FBNUI7QUFDQSxJQUFNRSxzQkFBc0IsR0FBR0YsTUFBTSxDQUFDLGdCQUFELENBQXJDO0FBQ0EsSUFBTUcsY0FBYyxHQUFHSCxNQUFNLENBQUMsU0FBRCxDQUE3QixDLENBRUE7QUFDQTtBQUNBOztJQUNlSSxzQjs7Ozs7RUFDYixrQ0FBbUM7SUFBQTs7SUFBQSxJQUF2QkMsT0FBdUIsdUVBQUosRUFBSTtJQUFBO0lBQ2pDO0lBRGlDLDJFQXlCbENOLDJCQXpCa0M7SUFBQSwyRUEwQmxDRSxhQTFCa0M7SUFBQSwyRUEyQmxDQyxzQkEzQmtDO0lBQUEsMkVBNEJsQ0MsY0E1QmtDO0lBR2pDLElBQU1HLGlCQUFvQyxHQUFHO01BQzNDQyxLQUFLLEVBQUVGLE9BQU8sQ0FBQ0UsS0FBUixJQUFpQixLQURtQjtNQUUzQ0MsRUFBRSxFQUFFSCxPQUFPLENBQUNHLEVBQVIsSUFBYyxJQUFBQyxRQUFBLElBQUtDLE9BQUwsQ0FBYSxPQUFiLEVBQW9CLEVBQXBCO0lBRnlCLENBQTdDLENBSGlDLENBUWpDO0lBQ0E7O0lBQ0EsTUFBS1gsMkJBQUwsSUFBb0MsSUFBQVksaUJBQUEsR0FBcEMsQ0FWaUMsQ0FZakM7SUFDQTs7SUFDQSxNQUFLVixhQUFMLElBQXNCLElBQUlXLG9CQUFKLEVBQXRCLENBZGlDLENBZ0JqQztJQUNBOztJQUNBLE1BQUtWLHNCQUFMLElBQStCLElBQUFTLGlCQUFBLEdBQS9CLENBbEJpQyxDQW9CakM7SUFDQTs7SUFDQSxNQUFLUixjQUFMLElBQXVCRyxpQkFBdkI7SUF0QmlDO0VBdUJsQzs7Ozs7SUFPRDtJQUNBO0lBQ0E7SUFDQSxlQUE0QztNQUMxQztNQUNBO01BQ0EsT0FBTyxLQUFLTCxhQUFMLENBQVA7SUFDRDtJQUVEO0lBQ0E7SUFDQTtJQUNBO0lBQ0E7SUFDQTtJQUNBO0lBQ0E7Ozs7U0FDQSxlQUE2QztNQUMzQyxLQUFLTSxLQUFMLENBQVcsbUJBQVgsRUFEMkMsQ0FHM0M7TUFDQTs7TUFDQSxPQUFPLEtBQUtMLHNCQUFMLEVBQTZCVyxPQUFwQztJQUNEO0lBRUQ7Ozs7V0FDQSxjQUFhO01BQ1g7TUFDQTtNQUNBLE9BQU8sS0FBS1YsY0FBTCxFQUFxQkssRUFBNUI7SUFDRDtJQUVEO0lBQ0E7SUFDQTs7OztXQUNBLGlCQUFrQk0sS0FBbEIsRUFBaUQ7TUFDL0M7TUFDQTtNQUNBLEtBQUtiLGFBQUwsRUFBb0JjLE9BQXBCLENBQTRCRCxLQUE1Qjs7TUFDQUUsZUFBQSxDQUFPQyxRQUFQLENBQWdCRixPQUFoQixDQUF3QkQsS0FBeEI7SUFDRDtJQUVEOzs7O1dBQ0EsNEJBQW1DO01BQ2pDLEtBQUtQLEtBQUwsQ0FBVywwQ0FBWDtNQUNBLEtBQUtRLE9BQUwsQ0FBYSxJQUFJRyxxQ0FBSixDQUFpQyxLQUFLVixFQUFMLEVBQWpDLENBQWI7SUFDRDtJQUVEOzs7O1dBQ0EscUJBQTRCO01BQzFCLEtBQUtELEtBQUwsQ0FBVyxtQ0FBWDtNQUNBLEtBQUtRLE9BQUwsQ0FBYSxJQUFJSSw4QkFBSixDQUEwQixLQUFLWCxFQUFMLEVBQTFCLENBQWI7SUFDRDtJQUVEO0lBQ0E7SUFDQTs7OztXQUNBLG1CQUFvQlksS0FBcEIsRUFBd0M7TUFDdEMsS0FBS2IsS0FBTCxDQUFXLG1DQUFYLEVBQWdEO1FBQUVhLEtBQUssRUFBTEE7TUFBRixDQUFoRCxFQURzQyxDQUd0Qzs7TUFDQSxLQUFLTCxPQUFMLENBQWEsSUFBSU0sOEJBQUosQ0FBMEIsS0FBS2IsRUFBTCxFQUExQixFQUFxQ1ksS0FBSyxDQUFDRSxPQUEzQyxDQUFiO0lBQ0Q7SUFFRDs7OztXQUNBLDJCQUE0QkMsV0FBNUIsRUFBdUQ7TUFDckQsS0FBS2hCLEtBQUwsaUVBQWlFZ0IsV0FBakU7TUFDQSxLQUFLUixPQUFMLENBQWEsSUFBSVMsc0NBQUosQ0FBa0MsS0FBS2hCLEVBQUwsRUFBbEMsRUFBNkNlLFdBQTdDLENBQWI7SUFDRDtJQUVEOzs7O1dBQ0EsMEJBQTJCQSxXQUEzQixFQUFzRDtNQUNwRCxLQUFLaEIsS0FBTCxnRUFBZ0VnQixXQUFoRTtNQUNBLEtBQUtSLE9BQUwsQ0FBYSxJQUFJVSxxQ0FBSixDQUFpQyxLQUFLakIsRUFBTCxFQUFqQyxFQUE0Q2UsV0FBNUMsQ0FBYjtJQUNEO0lBRUQ7Ozs7V0FDQSx1QkFBd0JBLFdBQXhCLEVBQTZDSCxLQUE3QyxFQUFpRTtNQUMvRCxLQUFLYixLQUFMLDZEQUE2RGdCLFdBQTdELFVBQThFO1FBQUVILEtBQUssRUFBTEE7TUFBRixDQUE5RSxFQUQrRCxDQUcvRDs7TUFDQSxLQUFLTCxPQUFMLENBQWEsSUFBSVcsa0NBQUosQ0FBOEIsS0FBS2xCLEVBQUwsRUFBOUIsRUFBeUNlLFdBQXpDLEVBQXNESCxLQUFLLENBQUNFLE9BQTVELENBQWI7SUFDRDtJQUVEOzs7O1dBQ0EsMEJBQTJCQyxXQUEzQixFQUFzRDtNQUNwRCxLQUFLaEIsS0FBTCxDQUFXLDBDQUFYO01BQ0EsS0FBS1EsT0FBTCxDQUFhLElBQUlZLHFDQUFKLENBQWlDLEtBQUtuQixFQUFMLEVBQWpDLEVBQTRDZSxXQUE1QyxDQUFiO0lBQ0Q7SUFFRDs7OztXQUNBLG1CQUEwQjtNQUN4QixLQUFLaEIsS0FBTCxDQUFXLGlDQUFYO01BQ0EsS0FBS1EsT0FBTCxDQUFhLElBQUlhLDRCQUFKLENBQXdCLEtBQUtwQixFQUFMLEVBQXhCLENBQWI7SUFDRCxDLENBRUQ7SUFDQTs7OztXQUNBLGlCQUFjO01BQ1osS0FBS0QsS0FBTCxDQUFXLHVCQUFYO01BRUEsTUFBTSxJQUFJc0IsS0FBSixDQUFVLGlCQUFWLENBQU47SUFDRCxDLENBRUQ7Ozs7V0FDQSxrQkFBZTtNQUNiLEtBQUt0QixLQUFMLENBQVcsd0JBQVg7TUFFQSxNQUFNLElBQUlzQixLQUFKLENBQVUsaUJBQVYsQ0FBTjtJQUNELEMsQ0FFRDs7OztXQUNBLGtCQUFlO01BQ2IsS0FBS3RCLEtBQUwsQ0FBVyx3QkFBWDtNQUVBLE1BQU0sSUFBSXNCLEtBQUosQ0FBVSxpQkFBVixDQUFOO0lBQ0Q7SUFFRDs7OztXQUNBLGVBQWNQLE9BQWQsRUFBZ0M7TUFBQTs7TUFBQSxrQ0FBTlEsSUFBTTtRQUFOQSxJQUFNO01BQUE7O01BQzlCO01BQ0E7TUFDQTtNQUNBLEtBQUszQixjQUFMLEVBQXFCSSxLQUFyQixJQUE4QixZQUFBd0IsT0FBTyxFQUFDQyxJQUFSLG9EQUF3Q1YsT0FBeEMsVUFBc0RRLElBQXRELEVBQTlCO0lBQ0Q7SUFFRDs7Ozs7SUFPQTtJQUNBLGdCQUFPUCxXQUFQLEVBQXNEO01BQUE7O01BQ3BELEtBQUtoQixLQUFMLDBDQUEwQ2dCLFdBQTFDO01BRUEsS0FBS1UsaUJBQUwsQ0FBdUJWLFdBQXZCO01BRUEsT0FBT1csT0FBTyxDQUFDQyxPQUFSLEdBQWtCQyxJQUFsQix1RkFBd0M7UUFBQTs7UUFBQTtVQUFBO1lBQUE7Y0FBQTtnQkFDN0MsTUFBSSxDQUFDQyxnQkFBTDs7Z0JBRDZDO2dCQUFBO2dCQUFBLE9BSVcsTUFBSSxDQUFDQyxhQUFMLENBQW1CZixXQUFuQixDQUpYOztjQUFBO2dCQUFBO2dCQUluQ2dCLGVBSm1DLHlCQUluQ0EsZUFKbUM7Z0JBSWxCQyxVQUprQix5QkFJbEJBLFVBSmtCO2dCQUlOQyxNQUpNLHlCQUlOQSxNQUpNOztnQkFNM0M7Z0JBQ0E7Z0JBRUE7Z0JBQ0E7Z0JBQ0EsTUFBSSxDQUFDMUMsMkJBQUQsQ0FBSixDQUFrQ29DLE9BQWxDLENBQTBDSyxVQUExQyxFQVgyQyxDQWEzQztnQkFDQTs7O2dCQUNBLE1BQUksQ0FBQ3RDLHNCQUFELENBQUosQ0FBNkJpQyxPQUE3QixDQUNFLElBQUlPLHdDQUFKLENBQTBCRCxNQUFNLENBQUNFLGFBQWpDLEVBQWdERixNQUFNLENBQUNHLGFBQXZELEVBQXNFSCxNQUFNLENBQUNJLFFBQTdFLENBREY7O2dCQUlBLE1BQUksQ0FBQ0MsU0FBTDs7Z0JBQ0EsTUFBSSxDQUFDQyxnQkFBTCxDQUFzQnhCLFdBQXRCOztnQkFwQjJDLGtDQXNCcEM7a0JBQ0x5QixNQUFNO29CQUFBLHNGQUFFO3NCQUFBO3dCQUFBOzBCQUFBOzRCQUFBOzhCQUNOLE1BQUksQ0FBQ3pDLEtBQUwsa0NBQW9DZ0IsV0FBcEM7OzhCQURNOzhCQUFBLE9BR0FnQixlQUFlLENBQUNTLE1BQWhCLEVBSEE7OzRCQUFBOzhCQUtOOzhCQUNBLE1BQUksQ0FBQ0MsT0FBTDs7OEJBQ0EsTUFBSSxDQUFDQyxnQkFBTCxDQUFzQjNCLFdBQXRCOzs0QkFQTTs0QkFBQTs4QkFBQTswQkFBQTt3QkFBQTtzQkFBQTtvQkFBQSxDQUFGOztvQkFBQTtzQkFBQTtvQkFBQTs7b0JBQUE7a0JBQUEsR0FERDtrQkFVTGYsRUFBRSxFQUFFO29CQUFBLE9BQU0rQixlQUFlLENBQUMvQixFQUFoQixFQUFOO2tCQUFBLENBVkM7a0JBV0wyQyxJQUFJLEVBQUUsZ0JBQU07b0JBQ1YsTUFBSSxDQUFDNUMsS0FBTCxDQUFXLFNBQVg7O29CQUVBLE9BQU9nQyxlQUFlLENBQUNZLElBQWhCLEVBQVA7a0JBQ0Q7Z0JBZkksQ0F0Qm9DOztjQUFBO2dCQUFBO2dCQUFBOztnQkF3QzNDLE1BQUksQ0FBQ0MsYUFBTCxDQUFtQjdCLFdBQW5COztnQkF4QzJDOztjQUFBO2NBQUE7Z0JBQUE7WUFBQTtVQUFBO1FBQUE7TUFBQSxDQUF4QyxHQUFQO0lBNkNEO0lBRUQ7QUFDRjtBQUNBO0FBQ0E7QUFDQTtJQUVFO0lBQ0E7Ozs7V0FDQSwwQkFBMEM7TUFDeEM7TUFDQTtNQUNBO0lBQ0Q7SUFFRDtJQUNBO0lBQ0E7SUFDQTtJQUNBOzs7Ozs2RkFDQTtRQUFBO1VBQUE7WUFBQTtjQUFBO2dCQUNFLEtBQUtoQixLQUFMO2dCQURGO2dCQUFBLE9BR1EsS0FBSzhDLGNBQUwsRUFIUjs7Y0FBQTtjQUFBO2dCQUFBO1lBQUE7VUFBQTtRQUFBO01BQUEsQzs7Ozs7Ozs7SUFNQTtJQUNBO0lBQ0E7Ozs7U0FDQSxlQUFvRDtNQUNsRCxLQUFLOUMsS0FBTCw0QkFEa0QsQ0FHbEQ7TUFDQTs7TUFDQSxPQUFPMkIsT0FBTyxDQUFDb0IsR0FBUixDQUFZLENBQUMsS0FBS3ZELDJCQUFMLEVBQWtDYyxPQUFuQyxFQUE0QyxLQUFLWCxzQkFBTCxFQUE2QlcsT0FBekUsQ0FBWixFQUErRnVCLElBQS9GLENBQ0w7UUFBQTtRQUFBO1FBQUEsSUFBSW1CLFlBQUosVUFBSUEsWUFBSjtRQUFBLElBQWtCQyxZQUFsQixVQUFrQkEsWUFBbEI7UUFBQSxJQUFnQ0MsS0FBaEMsVUFBZ0NBLEtBQWhDO1FBQUEsSUFBdUNDLElBQXZDLFVBQXVDQSxJQUF2QztRQUFBO1FBQUEsSUFBaURkLGFBQWpELFdBQWlEQSxhQUFqRDtRQUFBLElBQWdFQyxRQUFoRSxXQUFnRUEsUUFBaEU7UUFBQSxJQUEwRUYsYUFBMUUsV0FBMEVBLGFBQTFFOztRQUFBLE9BQWdHO1VBQzlGZ0IsYUFBYSxFQUFFZixhQUQrRTtVQUU5RmdCLFlBQVksRUFBRWYsUUFGZ0Y7VUFHOUZVLFlBQVksRUFDVixPQUFPQSxZQUFQLEtBQXdCLFFBQXhCLElBQW9DLENBQUMsSUFBQU0sZ0RBQUEsRUFBd0JOLFlBQXhCLENBQXJDLEdBQ0k7VUFDQTtVQUNBTyxzQkFBQSxDQUFhUCxZQUFiLENBSEosR0FJSUEsWUFBWSxJQUFJTyxzQkFBQSxDQUFhQyxPQVIyRDtVQVM5RlAsWUFBWSxFQUFFQSxZQUFZLElBQUksRUFUZ0U7VUFVOUZDLEtBQUssRUFBRUEsS0FBSyxJQUFJLEVBVjhFO1VBVzlGTyxVQUFVLEVBQUVyQixhQVhrRjtVQVk5RjtVQUNBO1VBQ0FlLElBQUksRUFBRSxPQUFPQSxJQUFQLEtBQWdCLFFBQWhCLElBQTRCLENBQUMsSUFBQUcsZ0RBQUEsRUFBd0JILElBQXhCLENBQTdCLEdBQTZETyxjQUFBLENBQUtQLElBQUwsQ0FBN0QsR0FBMEVBLElBQUksSUFBSU8sY0FBQSxDQUFLRjtRQWRDLENBQWhHO01BQUEsQ0FESyxDQUFQO0lBa0JEOzs7RUE1UTJDRyxxRDs7ZUErUS9COUQsc0IifQ==