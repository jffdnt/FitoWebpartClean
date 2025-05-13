"use strict";

var _interopRequireDefault = require("@babel/runtime/helpers/interopRequireDefault");

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.default = createMicrophoneAudioConfigAndAudioContext;

var _regenerator = _interopRequireDefault(require("@babel/runtime/regenerator"));

var _slicedToArray2 = _interopRequireDefault(require("@babel/runtime/helpers/slicedToArray"));

var _asyncToGenerator2 = _interopRequireDefault(require("@babel/runtime/helpers/asyncToGenerator"));

var _Exports = require("microsoft-cognitiveservices-speech-sdk/distrib/lib/src/common/Exports");

var _Exports2 = require("microsoft-cognitiveservices-speech-sdk/distrib/lib/src/common.browser/Exports");

var _bytesPerSample = _interopRequireDefault(require("./bytesPerSample"));

var _createAudioConfig = _interopRequireDefault(require("./createAudioConfig"));

var _createAudioContext = _interopRequireDefault(require("./createAudioContext"));

var _getUserMedia = _interopRequireDefault(require("./getUserMedia"));

// This is how often we are flushing audio buffer to the network. Modify this value will affect latency.
var DEFAULT_BUFFER_DURATION_IN_MS = 100; // TODO: [P2] #3975 We should consider building our own PcmRecorder:
//       - Use Audio Worklet via blob URL
//       - Not hardcoding the sample rate or other values
// PcmRecorder always downscale to 16000 Hz. We cannot use the dynamic value from MediaConstraints or MediaTrackSettings.

var PCM_RECORDER_HARDCODED_SETTINGS = Object.freeze({
  channelCount: 1,
  sampleRate: 16000,
  sampleSize: 16
});
var PCM_RECORDER_HARDCODED_FORMAT = Object.freeze({
  bitsPerSample: PCM_RECORDER_HARDCODED_SETTINGS.sampleSize,
  // `channelCount` is not on @types/web@0.0.54 yet, related to https://github.com/microsoft/TypeScript-DOM-lib-generator/issues/1290.
  // @ts-ignore
  channels: PCM_RECORDER_HARDCODED_SETTINGS.channelCount,
  samplesPerSec: PCM_RECORDER_HARDCODED_SETTINGS.sampleRate
});

function createMicrophoneAudioConfig(options) {
  var audioConstraints = options.audioConstraints,
      audioContext = options.audioContext,
      debug = options.debug,
      enableTelemetry = options.enableTelemetry,
      pcmRecorderWorkletUrl = options.pcmRecorderWorkletUrl;
  var bufferDurationInMS = options.bufferDurationInMS || DEFAULT_BUFFER_DURATION_IN_MS; // Related to #4523.
  // When bumping to recent version of `microsoft-cognitiveservices-speech-sdk@>=1.23.0`, pass `true` to the constructor.
  // const pcmRecorder = new PcmRecorder(true);

  var pcmRecorder = new _Exports2.PcmRecorder();
  pcmRecorderWorkletUrl && pcmRecorder.setWorkletUrl(pcmRecorderWorkletUrl);
  return (0, _createAudioConfig.default)({
    attach: function attach(audioNodeId) {
      return (0, _asyncToGenerator2.default)( /*#__PURE__*/_regenerator.default.mark(function _callee() {
        var mediaStream, _mediaStream$getAudio, _mediaStream$getAudio2, firstAudioTrack, outputStream;

        return _regenerator.default.wrap(function _callee$(_context) {
          while (1) {
            switch (_context.prev = _context.next) {
              case 0:
                _context.next = 2;
                return (0, _getUserMedia.default)({
                  audio: audioConstraints,
                  video: false
                });

              case 2:
                mediaStream = _context.sent;
                _mediaStream$getAudio = mediaStream.getAudioTracks(), _mediaStream$getAudio2 = (0, _slicedToArray2.default)(_mediaStream$getAudio, 1), firstAudioTrack = _mediaStream$getAudio2[0];

                if (firstAudioTrack) {
                  _context.next = 6;
                  break;
                }

                throw new Error('No audio device is found.');

              case 6:
                outputStream = new _Exports.ChunkedArrayBufferStream( // Speech SDK quirks: PcmRecorder hardcoded sample rate of 16000 Hz.
                (0, _bytesPerSample.default)(PCM_RECORDER_HARDCODED_SETTINGS) * ( // eslint-disable-next-line no-magic-numbers
                (bufferDurationInMS || DEFAULT_BUFFER_DURATION_IN_MS) / 1000), audioNodeId);
                pcmRecorder.record(audioContext, mediaStream, outputStream);
                return _context.abrupt("return", {
                  audioStreamNode: {
                    // Speech SDK quirks: In SDK's original MicAudioSource implementation, it call turnOff() during detach().
                    //                    That means, it call turnOff(), then detach(), then turnOff() again. Seems redundant.
                    //                    When using with Direct Line Speech, turnOff() is never called.
                    detach: function detach() {
                      // Speech SDK quirks: In SDK, it call outputStream.close() in turnOff() before outputStream.readEnded() in detach().
                      //                    I think it make sense to call readEnded() before close().
                      outputStream.readEnded();
                      outputStream.close(); // PcmRecorder.releaseMediaResources() will disconnect/stop the MediaStream.
                      // We cannot use MediaStream again after turned off.

                      pcmRecorder.releaseMediaResources(audioContext); // MediaStream will become inactive after all tracks are removed.

                      mediaStream.getTracks().forEach(function (track) {
                        return mediaStream.removeTrack(track);
                      }); // ESLint: "return" is required by TypeScript
                      // eslint-disable-next-line no-useless-return

                      return;
                    },
                    id: function id() {
                      return audioNodeId;
                    },
                    read: function read() {
                      return outputStream.read();
                    }
                  },
                  deviceInfo: {
                    manufacturer: 'Bot Framework Web Chat',
                    model: enableTelemetry ? firstAudioTrack.label : '',
                    type: enableTelemetry ? 'Microphones' : 'Unknown'
                  },
                  // Speech SDK quirks: PcmRecorder hardcoded sample rate of 16000 Hz.
                  //                    We cannot obtain this number other than looking at their source code.
                  //                    I.e. no getter property.
                  // PcmRecorder always downscale to 16000 Hz. We cannot use the dynamic value from MediaConstraints or MediaTrackSettings.
                  format: PCM_RECORDER_HARDCODED_FORMAT
                });

              case 9:
              case "end":
                return _context.stop();
            }
          }
        }, _callee);
      }))();
    },
    debug: debug
  });
}

function createMicrophoneAudioConfigAndAudioContext(_ref) {
  var audioContext = _ref.audioContext,
      audioInputDeviceId = _ref.audioInputDeviceId,
      enableTelemetry = _ref.enableTelemetry;
  // Web Chat has an implementation of AudioConfig for microphone that would enable better support on Safari:
  // - Maintain same instance of `AudioContext` across recognitions;
  // - Resume suspended `AudioContext` on user gestures.
  //
  // This is filed as https://github.com/microsoft/cognitive-services-speech-sdk-js/issues/571.
  // Before Speech SDK team take our suggestion, we need to continue using a custom AudioConfig object to persist the blessing.
  audioContext || (audioContext = (0, _createAudioContext.default)());
  return {
    audioConfig: createMicrophoneAudioConfig({
      audioConstraints: audioInputDeviceId ? {
        deviceId: audioInputDeviceId
      } : true,
      audioContext: audioContext,
      enableTelemetry: enableTelemetry ? true : undefined
    }),
    audioContext: audioContext
  };
}
//# sourceMappingURL=data:application/json;charset=utf-8;base64,eyJ2ZXJzaW9uIjozLCJuYW1lcyI6WyJERUZBVUxUX0JVRkZFUl9EVVJBVElPTl9JTl9NUyIsIlBDTV9SRUNPUkRFUl9IQVJEQ09ERURfU0VUVElOR1MiLCJPYmplY3QiLCJmcmVlemUiLCJjaGFubmVsQ291bnQiLCJzYW1wbGVSYXRlIiwic2FtcGxlU2l6ZSIsIlBDTV9SRUNPUkRFUl9IQVJEQ09ERURfRk9STUFUIiwiYml0c1BlclNhbXBsZSIsImNoYW5uZWxzIiwic2FtcGxlc1BlclNlYyIsImNyZWF0ZU1pY3JvcGhvbmVBdWRpb0NvbmZpZyIsIm9wdGlvbnMiLCJhdWRpb0NvbnN0cmFpbnRzIiwiYXVkaW9Db250ZXh0IiwiZGVidWciLCJlbmFibGVUZWxlbWV0cnkiLCJwY21SZWNvcmRlcldvcmtsZXRVcmwiLCJidWZmZXJEdXJhdGlvbkluTVMiLCJwY21SZWNvcmRlciIsIlBjbVJlY29yZGVyIiwic2V0V29ya2xldFVybCIsImNyZWF0ZUF1ZGlvQ29uZmlnIiwiYXR0YWNoIiwiYXVkaW9Ob2RlSWQiLCJnZXRVc2VyTWVkaWEiLCJhdWRpbyIsInZpZGVvIiwibWVkaWFTdHJlYW0iLCJnZXRBdWRpb1RyYWNrcyIsImZpcnN0QXVkaW9UcmFjayIsIkVycm9yIiwib3V0cHV0U3RyZWFtIiwiQ2h1bmtlZEFycmF5QnVmZmVyU3RyZWFtIiwiYnl0ZXNQZXJTYW1wbGUiLCJyZWNvcmQiLCJhdWRpb1N0cmVhbU5vZGUiLCJkZXRhY2giLCJyZWFkRW5kZWQiLCJjbG9zZSIsInJlbGVhc2VNZWRpYVJlc291cmNlcyIsImdldFRyYWNrcyIsImZvckVhY2giLCJ0cmFjayIsInJlbW92ZVRyYWNrIiwiaWQiLCJyZWFkIiwiZGV2aWNlSW5mbyIsIm1hbnVmYWN0dXJlciIsIm1vZGVsIiwibGFiZWwiLCJ0eXBlIiwiZm9ybWF0IiwiY3JlYXRlTWljcm9waG9uZUF1ZGlvQ29uZmlnQW5kQXVkaW9Db250ZXh0IiwiYXVkaW9JbnB1dERldmljZUlkIiwiY3JlYXRlQXVkaW9Db250ZXh0IiwiYXVkaW9Db25maWciLCJkZXZpY2VJZCIsInVuZGVmaW5lZCJdLCJzb3VyY2VSb290IjoiYnVuZGxlOi8vLyIsInNvdXJjZXMiOlsiLi4vLi4vc3JjL3NwZWVjaC9jcmVhdGVNaWNyb3Bob25lQXVkaW9Db25maWdBbmRBdWRpb0NvbnRleHQudHMiXSwic291cmNlc0NvbnRlbnQiOlsiaW1wb3J0IHsgQ2h1bmtlZEFycmF5QnVmZmVyU3RyZWFtIH0gZnJvbSAnbWljcm9zb2Z0LWNvZ25pdGl2ZXNlcnZpY2VzLXNwZWVjaC1zZGsvZGlzdHJpYi9saWIvc3JjL2NvbW1vbi9FeHBvcnRzJztcbmltcG9ydCB7IFBjbVJlY29yZGVyIH0gZnJvbSAnbWljcm9zb2Z0LWNvZ25pdGl2ZXNlcnZpY2VzLXNwZWVjaC1zZGsvZGlzdHJpYi9saWIvc3JjL2NvbW1vbi5icm93c2VyL0V4cG9ydHMnO1xuXG5pbXBvcnQgeyBBdWRpb1N0cmVhbU5vZGUsIERldmljZUluZm8sIEZvcm1hdCB9IGZyb20gJy4vQ3VzdG9tQXVkaW9JbnB1dFN0cmVhbSc7XG5pbXBvcnQgYnl0ZXNQZXJTYW1wbGUgZnJvbSAnLi9ieXRlc1BlclNhbXBsZSc7XG5pbXBvcnQgY3JlYXRlQXVkaW9Db25maWcgZnJvbSAnLi9jcmVhdGVBdWRpb0NvbmZpZyc7XG5pbXBvcnQgY3JlYXRlQXVkaW9Db250ZXh0IGZyb20gJy4vY3JlYXRlQXVkaW9Db250ZXh0JztcbmltcG9ydCBnZXRVc2VyTWVkaWEgZnJvbSAnLi9nZXRVc2VyTWVkaWEnO1xuXG4vLyBUaGlzIGlzIGhvdyBvZnRlbiB3ZSBhcmUgZmx1c2hpbmcgYXVkaW8gYnVmZmVyIHRvIHRoZSBuZXR3b3JrLiBNb2RpZnkgdGhpcyB2YWx1ZSB3aWxsIGFmZmVjdCBsYXRlbmN5LlxuY29uc3QgREVGQVVMVF9CVUZGRVJfRFVSQVRJT05fSU5fTVMgPSAxMDA7XG5cbi8vIFRPRE86IFtQMl0gIzM5NzUgV2Ugc2hvdWxkIGNvbnNpZGVyIGJ1aWxkaW5nIG91ciBvd24gUGNtUmVjb3JkZXI6XG4vLyAgICAgICAtIFVzZSBBdWRpbyBXb3JrbGV0IHZpYSBibG9iIFVSTFxuLy8gICAgICAgLSBOb3QgaGFyZGNvZGluZyB0aGUgc2FtcGxlIHJhdGUgb3Igb3RoZXIgdmFsdWVzXG4vLyBQY21SZWNvcmRlciBhbHdheXMgZG93bnNjYWxlIHRvIDE2MDAwIEh6LiBXZSBjYW5ub3QgdXNlIHRoZSBkeW5hbWljIHZhbHVlIGZyb20gTWVkaWFDb25zdHJhaW50cyBvciBNZWRpYVRyYWNrU2V0dGluZ3MuXG5jb25zdCBQQ01fUkVDT1JERVJfSEFSRENPREVEX1NFVFRJTkdTOiBNZWRpYVRyYWNrU2V0dGluZ3MgPSBPYmplY3QuZnJlZXplKHtcbiAgY2hhbm5lbENvdW50OiAxLFxuICBzYW1wbGVSYXRlOiAxNjAwMCxcbiAgc2FtcGxlU2l6ZTogMTZcbn0pO1xuXG5jb25zdCBQQ01fUkVDT1JERVJfSEFSRENPREVEX0ZPUk1BVDogRm9ybWF0ID0gT2JqZWN0LmZyZWV6ZSh7XG4gIGJpdHNQZXJTYW1wbGU6IFBDTV9SRUNPUkRFUl9IQVJEQ09ERURfU0VUVElOR1Muc2FtcGxlU2l6ZSxcbiAgLy8gYGNoYW5uZWxDb3VudGAgaXMgbm90IG9uIEB0eXBlcy93ZWJAMC4wLjU0IHlldCwgcmVsYXRlZCB0byBodHRwczovL2dpdGh1Yi5jb20vbWljcm9zb2Z0L1R5cGVTY3JpcHQtRE9NLWxpYi1nZW5lcmF0b3IvaXNzdWVzLzEyOTAuXG4gIC8vIEB0cy1pZ25vcmVcbiAgY2hhbm5lbHM6IFBDTV9SRUNPUkRFUl9IQVJEQ09ERURfU0VUVElOR1MuY2hhbm5lbENvdW50LFxuICBzYW1wbGVzUGVyU2VjOiBQQ01fUkVDT1JERVJfSEFSRENPREVEX1NFVFRJTkdTLnNhbXBsZVJhdGVcbn0pO1xuXG50eXBlIE1pY3JvcGhvbmVBdWRpb0lucHV0U3RyZWFtT3B0aW9ucyA9IHtcbiAgLyoqIFNwZWNpZmllcyB0aGUgY29uc3RyYWludHMgZm9yIHNlbGVjdGluZyBhbiBhdWRpbyBkZXZpY2UuICovXG4gIGF1ZGlvQ29uc3RyYWludHM/OiB0cnVlIHwgTWVkaWFUcmFja0NvbnN0cmFpbnRzO1xuXG4gIC8qKiBTcGVjaWZpZXMgdGhlIGBBdWRpb0NvbnRleHRgIHRvIHVzZS4gVGhpcyBvYmplY3QgbXVzdCBiZSBwcmltZWQgYW5kIHJlYWR5IHRvIHVzZS4gKi9cbiAgYXVkaW9Db250ZXh0OiBBdWRpb0NvbnRleHQ7XG5cbiAgLyoqIFNwZWNpZmllcyB0aGUgYnVmZmVyaW5nIGRlbGF5IG9uIGhvdyBvZnRlbiB0byBmbHVzaCBhdWRpbyBkYXRhIHRvIG5ldHdvcmsuIEluY3JlYXNpbmcgdGhlIHZhbHVlIHdpbGwgaW5jcmVhc2UgYXVkaW8gbGF0ZW5jeS4gRGVmYXVsdCBpcyAxMDAgbXMuICovXG4gIGJ1ZmZlckR1cmF0aW9uSW5NUz86IG51bWJlcjtcblxuICAvKiogU3BlY2lmaWVzIHdoZXRoZXIgdG8gZGlzcGxheSBkaWFnbm9zdGljIGluZm9ybWF0aW9uLiAqL1xuICBkZWJ1Zz86IHRydWU7XG5cbiAgLyoqIFNwZWNpZmllcyBpZiB0ZWxlbWV0cnkgZGF0YSBzaG91bGQgYmUgc2VudC4gSWYgbm90IHNwZWNpZmllZCwgdGVsZW1ldHJ5IGRhdGEgd2lsbCBOT1QgYmUgc2VudC4gKi9cbiAgZW5hYmxlVGVsZW1ldHJ5PzogdHJ1ZTtcblxuICAvKiogU3BlY2lmaWVzIHRoZSBgQXVkaW9Xb3JrbGV0YCBVUkwgZm9yIGBQY21SZWNvcmRlcmAuIElmIG5vdCBzcGVjaWZpZWQsIHdpbGwgdXNlIHNjcmlwdCBwcm9jZXNzb3Igb24gVUkgdGhyZWFkIGluc3RlYWQuICovXG4gIHBjbVJlY29yZGVyV29ya2xldFVybD86IHN0cmluZztcbn07XG5cbmZ1bmN0aW9uIGNyZWF0ZU1pY3JvcGhvbmVBdWRpb0NvbmZpZyhvcHRpb25zOiBNaWNyb3Bob25lQXVkaW9JbnB1dFN0cmVhbU9wdGlvbnMpIHtcbiAgY29uc3QgeyBhdWRpb0NvbnN0cmFpbnRzLCBhdWRpb0NvbnRleHQsIGRlYnVnLCBlbmFibGVUZWxlbWV0cnksIHBjbVJlY29yZGVyV29ya2xldFVybCB9ID0gb3B0aW9ucztcbiAgY29uc3QgYnVmZmVyRHVyYXRpb25Jbk1TID0gb3B0aW9ucy5idWZmZXJEdXJhdGlvbkluTVMgfHwgREVGQVVMVF9CVUZGRVJfRFVSQVRJT05fSU5fTVM7XG5cbiAgLy8gUmVsYXRlZCB0byAjNDUyMy5cbiAgLy8gV2hlbiBidW1waW5nIHRvIHJlY2VudCB2ZXJzaW9uIG9mIGBtaWNyb3NvZnQtY29nbml0aXZlc2VydmljZXMtc3BlZWNoLXNka0A+PTEuMjMuMGAsIHBhc3MgYHRydWVgIHRvIHRoZSBjb25zdHJ1Y3Rvci5cbiAgLy8gY29uc3QgcGNtUmVjb3JkZXIgPSBuZXcgUGNtUmVjb3JkZXIodHJ1ZSk7XG4gIGNvbnN0IHBjbVJlY29yZGVyID0gbmV3IFBjbVJlY29yZGVyKCk7XG5cbiAgcGNtUmVjb3JkZXJXb3JrbGV0VXJsICYmIHBjbVJlY29yZGVyLnNldFdvcmtsZXRVcmwocGNtUmVjb3JkZXJXb3JrbGV0VXJsKTtcblxuICByZXR1cm4gY3JlYXRlQXVkaW9Db25maWcoe1xuICAgIGFzeW5jIGF0dGFjaChhdWRpb05vZGVJZDogc3RyaW5nKTogUHJvbWlzZTx7XG4gICAgICBhdWRpb1N0cmVhbU5vZGU6IEF1ZGlvU3RyZWFtTm9kZTtcbiAgICAgIGRldmljZUluZm86IERldmljZUluZm87XG4gICAgICBmb3JtYXQ6IEZvcm1hdDtcbiAgICB9PiB7XG4gICAgICAvLyBXZSBuZWVkIHRvIGdldCBuZXcgTWVkaWFTdHJlYW0gb24gZXZlcnkgYXR0YWNoKCkuXG4gICAgICAvLyBUaGlzIGlzIGJlY2F1c2UgUGNtUmVjb3JkZXIucmVsZWFzZU1lZGlhUmVzb3VyY2VzKCkgZGlzY29ubmVjdGVkL3N0b3BwZWQgdGhlbS5cbiAgICAgIGNvbnN0IG1lZGlhU3RyZWFtID0gYXdhaXQgZ2V0VXNlck1lZGlhKHsgYXVkaW86IGF1ZGlvQ29uc3RyYWludHMsIHZpZGVvOiBmYWxzZSB9KTtcblxuICAgICAgY29uc3QgW2ZpcnN0QXVkaW9UcmFja10gPSBtZWRpYVN0cmVhbS5nZXRBdWRpb1RyYWNrcygpO1xuXG4gICAgICBpZiAoIWZpcnN0QXVkaW9UcmFjaykge1xuICAgICAgICB0aHJvdyBuZXcgRXJyb3IoJ05vIGF1ZGlvIGRldmljZSBpcyBmb3VuZC4nKTtcbiAgICAgIH1cblxuICAgICAgY29uc3Qgb3V0cHV0U3RyZWFtID0gbmV3IENodW5rZWRBcnJheUJ1ZmZlclN0cmVhbShcbiAgICAgICAgLy8gU3BlZWNoIFNESyBxdWlya3M6IFBjbVJlY29yZGVyIGhhcmRjb2RlZCBzYW1wbGUgcmF0ZSBvZiAxNjAwMCBIei5cbiAgICAgICAgYnl0ZXNQZXJTYW1wbGUoUENNX1JFQ09SREVSX0hBUkRDT0RFRF9TRVRUSU5HUykgKlxuICAgICAgICAgIC8vIGVzbGludC1kaXNhYmxlLW5leHQtbGluZSBuby1tYWdpYy1udW1iZXJzXG4gICAgICAgICAgKChidWZmZXJEdXJhdGlvbkluTVMgfHwgREVGQVVMVF9CVUZGRVJfRFVSQVRJT05fSU5fTVMpIC8gMTAwMCksXG4gICAgICAgIGF1ZGlvTm9kZUlkXG4gICAgICApO1xuXG4gICAgICBwY21SZWNvcmRlci5yZWNvcmQoYXVkaW9Db250ZXh0LCBtZWRpYVN0cmVhbSwgb3V0cHV0U3RyZWFtKTtcblxuICAgICAgcmV0dXJuIHtcbiAgICAgICAgYXVkaW9TdHJlYW1Ob2RlOiB7XG4gICAgICAgICAgLy8gU3BlZWNoIFNESyBxdWlya3M6IEluIFNESydzIG9yaWdpbmFsIE1pY0F1ZGlvU291cmNlIGltcGxlbWVudGF0aW9uLCBpdCBjYWxsIHR1cm5PZmYoKSBkdXJpbmcgZGV0YWNoKCkuXG4gICAgICAgICAgLy8gICAgICAgICAgICAgICAgICAgIFRoYXQgbWVhbnMsIGl0IGNhbGwgdHVybk9mZigpLCB0aGVuIGRldGFjaCgpLCB0aGVuIHR1cm5PZmYoKSBhZ2Fpbi4gU2VlbXMgcmVkdW5kYW50LlxuICAgICAgICAgIC8vICAgICAgICAgICAgICAgICAgICBXaGVuIHVzaW5nIHdpdGggRGlyZWN0IExpbmUgU3BlZWNoLCB0dXJuT2ZmKCkgaXMgbmV2ZXIgY2FsbGVkLlxuICAgICAgICAgIGRldGFjaDogKCk6IFByb21pc2U8dm9pZD4gPT4ge1xuICAgICAgICAgICAgLy8gU3BlZWNoIFNESyBxdWlya3M6IEluIFNESywgaXQgY2FsbCBvdXRwdXRTdHJlYW0uY2xvc2UoKSBpbiB0dXJuT2ZmKCkgYmVmb3JlIG91dHB1dFN0cmVhbS5yZWFkRW5kZWQoKSBpbiBkZXRhY2goKS5cbiAgICAgICAgICAgIC8vICAgICAgICAgICAgICAgICAgICBJIHRoaW5rIGl0IG1ha2Ugc2Vuc2UgdG8gY2FsbCByZWFkRW5kZWQoKSBiZWZvcmUgY2xvc2UoKS5cbiAgICAgICAgICAgIG91dHB1dFN0cmVhbS5yZWFkRW5kZWQoKTtcbiAgICAgICAgICAgIG91dHB1dFN0cmVhbS5jbG9zZSgpO1xuXG4gICAgICAgICAgICAvLyBQY21SZWNvcmRlci5yZWxlYXNlTWVkaWFSZXNvdXJjZXMoKSB3aWxsIGRpc2Nvbm5lY3Qvc3RvcCB0aGUgTWVkaWFTdHJlYW0uXG4gICAgICAgICAgICAvLyBXZSBjYW5ub3QgdXNlIE1lZGlhU3RyZWFtIGFnYWluIGFmdGVyIHR1cm5lZCBvZmYuXG4gICAgICAgICAgICBwY21SZWNvcmRlci5yZWxlYXNlTWVkaWFSZXNvdXJjZXMoYXVkaW9Db250ZXh0KTtcblxuICAgICAgICAgICAgLy8gTWVkaWFTdHJlYW0gd2lsbCBiZWNvbWUgaW5hY3RpdmUgYWZ0ZXIgYWxsIHRyYWNrcyBhcmUgcmVtb3ZlZC5cbiAgICAgICAgICAgIG1lZGlhU3RyZWFtLmdldFRyYWNrcygpLmZvckVhY2godHJhY2sgPT4gbWVkaWFTdHJlYW0ucmVtb3ZlVHJhY2sodHJhY2spKTtcblxuICAgICAgICAgICAgLy8gRVNMaW50OiBcInJldHVyblwiIGlzIHJlcXVpcmVkIGJ5IFR5cGVTY3JpcHRcbiAgICAgICAgICAgIC8vIGVzbGludC1kaXNhYmxlLW5leHQtbGluZSBuby11c2VsZXNzLXJldHVyblxuICAgICAgICAgICAgcmV0dXJuO1xuICAgICAgICAgIH0sXG4gICAgICAgICAgaWQ6ICgpID0+IGF1ZGlvTm9kZUlkLFxuICAgICAgICAgIHJlYWQ6ICgpID0+IG91dHB1dFN0cmVhbS5yZWFkKClcbiAgICAgICAgfSxcbiAgICAgICAgZGV2aWNlSW5mbzoge1xuICAgICAgICAgIG1hbnVmYWN0dXJlcjogJ0JvdCBGcmFtZXdvcmsgV2ViIENoYXQnLFxuICAgICAgICAgIG1vZGVsOiBlbmFibGVUZWxlbWV0cnkgPyBmaXJzdEF1ZGlvVHJhY2subGFiZWwgOiAnJyxcbiAgICAgICAgICB0eXBlOiBlbmFibGVUZWxlbWV0cnkgPyAnTWljcm9waG9uZXMnIDogJ1Vua25vd24nXG4gICAgICAgIH0sXG4gICAgICAgIC8vIFNwZWVjaCBTREsgcXVpcmtzOiBQY21SZWNvcmRlciBoYXJkY29kZWQgc2FtcGxlIHJhdGUgb2YgMTYwMDAgSHouXG4gICAgICAgIC8vICAgICAgICAgICAgICAgICAgICBXZSBjYW5ub3Qgb2J0YWluIHRoaXMgbnVtYmVyIG90aGVyIHRoYW4gbG9va2luZyBhdCB0aGVpciBzb3VyY2UgY29kZS5cbiAgICAgICAgLy8gICAgICAgICAgICAgICAgICAgIEkuZS4gbm8gZ2V0dGVyIHByb3BlcnR5LlxuICAgICAgICAvLyBQY21SZWNvcmRlciBhbHdheXMgZG93bnNjYWxlIHRvIDE2MDAwIEh6LiBXZSBjYW5ub3QgdXNlIHRoZSBkeW5hbWljIHZhbHVlIGZyb20gTWVkaWFDb25zdHJhaW50cyBvciBNZWRpYVRyYWNrU2V0dGluZ3MuXG4gICAgICAgIGZvcm1hdDogUENNX1JFQ09SREVSX0hBUkRDT0RFRF9GT1JNQVRcbiAgICAgIH07XG4gICAgfSxcbiAgICBkZWJ1Z1xuICB9KTtcbn1cblxuZXhwb3J0IGRlZmF1bHQgZnVuY3Rpb24gY3JlYXRlTWljcm9waG9uZUF1ZGlvQ29uZmlnQW5kQXVkaW9Db250ZXh0KHtcbiAgYXVkaW9Db250ZXh0LFxuICBhdWRpb0lucHV0RGV2aWNlSWQsXG4gIGVuYWJsZVRlbGVtZXRyeVxufToge1xuICBhdWRpb0NvbnRleHQ/OiBBdWRpb0NvbnRleHQ7XG4gIGF1ZGlvSW5wdXREZXZpY2VJZD86IHN0cmluZztcbiAgZW5hYmxlVGVsZW1ldHJ5PzogdHJ1ZTtcbn0pIHtcbiAgLy8gV2ViIENoYXQgaGFzIGFuIGltcGxlbWVudGF0aW9uIG9mIEF1ZGlvQ29uZmlnIGZvciBtaWNyb3Bob25lIHRoYXQgd291bGQgZW5hYmxlIGJldHRlciBzdXBwb3J0IG9uIFNhZmFyaTpcbiAgLy8gLSBNYWludGFpbiBzYW1lIGluc3RhbmNlIG9mIGBBdWRpb0NvbnRleHRgIGFjcm9zcyByZWNvZ25pdGlvbnM7XG4gIC8vIC0gUmVzdW1lIHN1c3BlbmRlZCBgQXVkaW9Db250ZXh0YCBvbiB1c2VyIGdlc3R1cmVzLlxuICAvL1xuICAvLyBUaGlzIGlzIGZpbGVkIGFzIGh0dHBzOi8vZ2l0aHViLmNvbS9taWNyb3NvZnQvY29nbml0aXZlLXNlcnZpY2VzLXNwZWVjaC1zZGstanMvaXNzdWVzLzU3MS5cbiAgLy8gQmVmb3JlIFNwZWVjaCBTREsgdGVhbSB0YWtlIG91ciBzdWdnZXN0aW9uLCB3ZSBuZWVkIHRvIGNvbnRpbnVlIHVzaW5nIGEgY3VzdG9tIEF1ZGlvQ29uZmlnIG9iamVjdCB0byBwZXJzaXN0IHRoZSBibGVzc2luZy5cbiAgYXVkaW9Db250ZXh0IHx8IChhdWRpb0NvbnRleHQgPSBjcmVhdGVBdWRpb0NvbnRleHQoKSk7XG5cbiAgcmV0dXJuIHtcbiAgICBhdWRpb0NvbmZpZzogY3JlYXRlTWljcm9waG9uZUF1ZGlvQ29uZmlnKHtcbiAgICAgIGF1ZGlvQ29uc3RyYWludHM6IGF1ZGlvSW5wdXREZXZpY2VJZCA/IHsgZGV2aWNlSWQ6IGF1ZGlvSW5wdXREZXZpY2VJZCB9IDogdHJ1ZSxcbiAgICAgIGF1ZGlvQ29udGV4dCxcbiAgICAgIGVuYWJsZVRlbGVtZXRyeTogZW5hYmxlVGVsZW1ldHJ5ID8gdHJ1ZSA6IHVuZGVmaW5lZFxuICAgIH0pLFxuICAgIGF1ZGlvQ29udGV4dFxuICB9O1xufVxuIl0sIm1hcHBpbmdzIjoiOzs7Ozs7Ozs7Ozs7Ozs7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