"use strict";

var _interopRequireDefault = require("@babel/runtime/helpers/interopRequireDefault");

Object.defineProperty(exports, "__esModule", {
  value: true
});
var _exportNames = {
  createDirectLine: true,
  createDirectLineAppServiceExtension: true
};
exports.createDirectLineAppServiceExtension = exports.createDirectLine = void 0;

var _defineProperty2 = _interopRequireDefault(require("@babel/runtime/helpers/defineProperty"));

require("./polyfill");

var _indexMinimal = require("./index-minimal");

var _addVersion = _interopRequireDefault(require("./addVersion"));

var _createDirectLine = _interopRequireDefault(require("./createDirectLine"));

var _createDirectLineAppServiceExtension = _interopRequireDefault(require("./createDirectLineAppServiceExtension"));

var _index = require("./index");

Object.keys(_index).forEach(function (key) {
  if (key === "default" || key === "__esModule") return;
  if (Object.prototype.hasOwnProperty.call(_exportNames, key)) return;
  if (key in exports && exports[key] === _index[key]) return;
  Object.defineProperty(exports, key, {
    enumerable: true,
    get: function get() {
      return _index[key];
    }
  });
});

function ownKeys(object, enumerableOnly) { var keys = Object.keys(object); if (Object.getOwnPropertySymbols) { var symbols = Object.getOwnPropertySymbols(object); enumerableOnly && (symbols = symbols.filter(function (sym) { return Object.getOwnPropertyDescriptor(object, sym).enumerable; })), keys.push.apply(keys, symbols); } return keys; }

function _objectSpread(target) { for (var i = 1; i < arguments.length; i++) { var source = null != arguments[i] ? arguments[i] : {}; i % 2 ? ownKeys(Object(source), !0).forEach(function (key) { (0, _defineProperty2.default)(target, key, source[key]); }) : Object.getOwnPropertyDescriptors ? Object.defineProperties(target, Object.getOwnPropertyDescriptors(source)) : ownKeys(Object(source)).forEach(function (key) { Object.defineProperty(target, key, Object.getOwnPropertyDescriptor(source, key)); }); } return target; }

var createDirectLine = function createDirectLine(options) {
  options.botAgent && console.warn('Web Chat: Developers are not currently allowed to set botAgent in the createDirectLine function. See https://github.com/microsoft/BotFramework-WebChat/issues/2119 for more details.');
  return (0, _createDirectLine.default)(_objectSpread(_objectSpread({}, options), {}, {
    botAgent: "WebChat/".concat(_indexMinimal.version, " (ES5)")
  }));
};

exports.createDirectLine = createDirectLine;

var createDirectLineAppServiceExtension = function createDirectLineAppServiceExtension(options) {
  options.botAgent && console.warn('Web Chat: Developers are not currently allowed to set botAgent in the createDirectLine function. See https://github.com/microsoft/BotFramework-WebChat/issues/2119 for more details.');
  return (0, _createDirectLineAppServiceExtension.default)(_objectSpread(_objectSpread({}, options), {}, {
    botAgent: "WebChat/".concat(_indexMinimal.version, " (ES5)")
  }));
};

exports.createDirectLineAppServiceExtension = createDirectLineAppServiceExtension;
window['WebChat'] = _objectSpread(_objectSpread({}, window['WebChat']), {}, {
  createDirectLine: createDirectLine,
  createDirectLineAppServiceExtension: createDirectLineAppServiceExtension
});
(0, _addVersion.default)('full-es5');
//# sourceMappingURL=data:application/json;charset=utf-8;base64,eyJ2ZXJzaW9uIjozLCJuYW1lcyI6WyJjcmVhdGVEaXJlY3RMaW5lIiwib3B0aW9ucyIsImJvdEFnZW50IiwiY29uc29sZSIsIndhcm4iLCJkZWZhdWx0Q3JlYXRlRGlyZWN0TGluZSIsInZlcnNpb24iLCJjcmVhdGVEaXJlY3RMaW5lQXBwU2VydmljZUV4dGVuc2lvbiIsImRlZmF1bHRDcmVhdGVEaXJlY3RMaW5lQXBwU2VydmljZUV4dGVuc2lvbiIsIndpbmRvdyIsImFkZFZlcnNpb24iXSwic291cmNlUm9vdCI6ImJ1bmRsZTovLy8iLCJzb3VyY2VzIjpbIi4uL3NyYy9pbmRleC1lczUudHMiXSwic291cmNlc0NvbnRlbnQiOlsiLyogZXNsaW50IGRvdC1ub3RhdGlvbjogW1wiZXJyb3JcIiwgeyBcImFsbG93UGF0dGVyblwiOiBcIl5XZWJDaGF0JFwiIH1dICovXG4vLyB3aW5kb3dbJ1dlYkNoYXQnXSBpcyByZXF1aXJlZCBmb3IgVHlwZVNjcmlwdFxuXG4vLyBJbXBvcnRpbmcgcG9seWZpbGxzIHJlcXVpcmVkIGZvciBJRTExL0VTNS5cbmltcG9ydCAnLi9wb2x5ZmlsbCc7XG5cbmltcG9ydCB7IHZlcnNpb24gfSBmcm9tICcuL2luZGV4LW1pbmltYWwnO1xuaW1wb3J0IGFkZFZlcnNpb24gZnJvbSAnLi9hZGRWZXJzaW9uJztcbmltcG9ydCBkZWZhdWx0Q3JlYXRlRGlyZWN0TGluZSBmcm9tICcuL2NyZWF0ZURpcmVjdExpbmUnO1xuaW1wb3J0IGRlZmF1bHRDcmVhdGVEaXJlY3RMaW5lQXBwU2VydmljZUV4dGVuc2lvbiBmcm9tICcuL2NyZWF0ZURpcmVjdExpbmVBcHBTZXJ2aWNlRXh0ZW5zaW9uJztcblxuZXhwb3J0ICogZnJvbSAnLi9pbmRleCc7XG5cbmV4cG9ydCBjb25zdCBjcmVhdGVEaXJlY3RMaW5lID0gb3B0aW9ucyA9PiB7XG4gIG9wdGlvbnMuYm90QWdlbnQgJiZcbiAgICBjb25zb2xlLndhcm4oXG4gICAgICAnV2ViIENoYXQ6IERldmVsb3BlcnMgYXJlIG5vdCBjdXJyZW50bHkgYWxsb3dlZCB0byBzZXQgYm90QWdlbnQgaW4gdGhlIGNyZWF0ZURpcmVjdExpbmUgZnVuY3Rpb24uIFNlZSBodHRwczovL2dpdGh1Yi5jb20vbWljcm9zb2Z0L0JvdEZyYW1ld29yay1XZWJDaGF0L2lzc3Vlcy8yMTE5IGZvciBtb3JlIGRldGFpbHMuJ1xuICAgICk7XG5cbiAgcmV0dXJuIGRlZmF1bHRDcmVhdGVEaXJlY3RMaW5lKHsgLi4ub3B0aW9ucywgYm90QWdlbnQ6IGBXZWJDaGF0LyR7dmVyc2lvbn0gKEVTNSlgIH0pO1xufTtcblxuZXhwb3J0IGNvbnN0IGNyZWF0ZURpcmVjdExpbmVBcHBTZXJ2aWNlRXh0ZW5zaW9uID0gb3B0aW9ucyA9PiB7XG4gIG9wdGlvbnMuYm90QWdlbnQgJiZcbiAgICBjb25zb2xlLndhcm4oXG4gICAgICAnV2ViIENoYXQ6IERldmVsb3BlcnMgYXJlIG5vdCBjdXJyZW50bHkgYWxsb3dlZCB0byBzZXQgYm90QWdlbnQgaW4gdGhlIGNyZWF0ZURpcmVjdExpbmUgZnVuY3Rpb24uIFNlZSBodHRwczovL2dpdGh1Yi5jb20vbWljcm9zb2Z0L0JvdEZyYW1ld29yay1XZWJDaGF0L2lzc3Vlcy8yMTE5IGZvciBtb3JlIGRldGFpbHMuJ1xuICAgICk7XG5cbiAgcmV0dXJuIGRlZmF1bHRDcmVhdGVEaXJlY3RMaW5lQXBwU2VydmljZUV4dGVuc2lvbih7IC4uLm9wdGlvbnMsIGJvdEFnZW50OiBgV2ViQ2hhdC8ke3ZlcnNpb259IChFUzUpYCB9KTtcbn07XG5cbndpbmRvd1snV2ViQ2hhdCddID0ge1xuICAuLi53aW5kb3dbJ1dlYkNoYXQnXSxcbiAgY3JlYXRlRGlyZWN0TGluZSxcbiAgY3JlYXRlRGlyZWN0TGluZUFwcFNlcnZpY2VFeHRlbnNpb25cbn07XG5cbmFkZFZlcnNpb24oJ2Z1bGwtZXM1Jyk7XG4iXSwibWFwcGluZ3MiOiI7Ozs7Ozs7Ozs7Ozs7OztBQUlBOztBQUVBOztBQUNBOztBQUNBOztBQUNBOztBQUVBOztBQUFBO0VBQUE7RUFBQTtFQUFBO0VBQUE7SUFBQTtJQUFBO01BQUE7SUFBQTtFQUFBO0FBQUE7Ozs7OztBQUVPLElBQU1BLGdCQUFnQixHQUFHLFNBQW5CQSxnQkFBbUIsQ0FBQUMsT0FBTyxFQUFJO0VBQ3pDQSxPQUFPLENBQUNDLFFBQVIsSUFDRUMsT0FBTyxDQUFDQyxJQUFSLENBQ0Usc0xBREYsQ0FERjtFQUtBLE9BQU8sSUFBQUMseUJBQUEsa0NBQTZCSixPQUE3QjtJQUFzQ0MsUUFBUSxvQkFBYUkscUJBQWI7RUFBOUMsR0FBUDtBQUNELENBUE07Ozs7QUFTQSxJQUFNQyxtQ0FBbUMsR0FBRyxTQUF0Q0EsbUNBQXNDLENBQUFOLE9BQU8sRUFBSTtFQUM1REEsT0FBTyxDQUFDQyxRQUFSLElBQ0VDLE9BQU8sQ0FBQ0MsSUFBUixDQUNFLHNMQURGLENBREY7RUFLQSxPQUFPLElBQUFJLDRDQUFBLGtDQUFnRFAsT0FBaEQ7SUFBeURDLFFBQVEsb0JBQWFJLHFCQUFiO0VBQWpFLEdBQVA7QUFDRCxDQVBNOzs7QUFTUEcsTUFBTSxDQUFDLFNBQUQsQ0FBTixtQ0FDS0EsTUFBTSxDQUFDLFNBQUQsQ0FEWDtFQUVFVCxnQkFBZ0IsRUFBaEJBLGdCQUZGO0VBR0VPLG1DQUFtQyxFQUFuQ0E7QUFIRjtBQU1BLElBQUFHLG1CQUFBLEVBQVcsVUFBWCJ9