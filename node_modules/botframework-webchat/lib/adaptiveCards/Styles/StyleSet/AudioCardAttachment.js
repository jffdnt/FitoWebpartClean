"use strict";

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.default = AudioCardAttachment;

// TODO: [P4] We are moving attachments related to Adaptive Cards out of "component"
//       Later, we will rewrite these attachments without Adaptive Cards
//       We are leaving the CSS here as-is for now
function AudioCardAttachment(_ref) {
  var _ref$paddingRegular = _ref.paddingRegular,
      paddingRegular = _ref$paddingRegular === void 0 ? undefined : _ref$paddingRegular;
  return {
    // TODO: [P2] We should not set "display" in styleSet, this will allow the user to break the layout for no good reasons.
    display: 'flex',
    flexDirection: 'column',
    '& > ul.media-list': {
      // TODO: [P2] We should not set "listStyleType" in styleSet, the user have no good reasons to change this.
      listStyleType: 'none',
      margin: paddingRegular,
      padding: 0
    }
  };
}
//# sourceMappingURL=data:application/json;charset=utf-8;base64,eyJ2ZXJzaW9uIjozLCJuYW1lcyI6WyJBdWRpb0NhcmRBdHRhY2htZW50IiwicGFkZGluZ1JlZ3VsYXIiLCJ1bmRlZmluZWQiLCJkaXNwbGF5IiwiZmxleERpcmVjdGlvbiIsImxpc3RTdHlsZVR5cGUiLCJtYXJnaW4iLCJwYWRkaW5nIl0sInNvdXJjZVJvb3QiOiJidW5kbGU6Ly8vIiwic291cmNlcyI6WyIuLi8uLi8uLi8uLi9zcmMvYWRhcHRpdmVDYXJkcy9TdHlsZXMvU3R5bGVTZXQvQXVkaW9DYXJkQXR0YWNobWVudC5qcyJdLCJzb3VyY2VzQ29udGVudCI6WyIvLyBUT0RPOiBbUDRdIFdlIGFyZSBtb3ZpbmcgYXR0YWNobWVudHMgcmVsYXRlZCB0byBBZGFwdGl2ZSBDYXJkcyBvdXQgb2YgXCJjb21wb25lbnRcIlxuLy8gICAgICAgTGF0ZXIsIHdlIHdpbGwgcmV3cml0ZSB0aGVzZSBhdHRhY2htZW50cyB3aXRob3V0IEFkYXB0aXZlIENhcmRzXG4vLyAgICAgICBXZSBhcmUgbGVhdmluZyB0aGUgQ1NTIGhlcmUgYXMtaXMgZm9yIG5vd1xuXG5leHBvcnQgZGVmYXVsdCBmdW5jdGlvbiBBdWRpb0NhcmRBdHRhY2htZW50KHsgcGFkZGluZ1JlZ3VsYXIgPSB1bmRlZmluZWQgfSkge1xuICByZXR1cm4ge1xuICAgIC8vIFRPRE86IFtQMl0gV2Ugc2hvdWxkIG5vdCBzZXQgXCJkaXNwbGF5XCIgaW4gc3R5bGVTZXQsIHRoaXMgd2lsbCBhbGxvdyB0aGUgdXNlciB0byBicmVhayB0aGUgbGF5b3V0IGZvciBubyBnb29kIHJlYXNvbnMuXG4gICAgZGlzcGxheTogJ2ZsZXgnLFxuICAgIGZsZXhEaXJlY3Rpb246ICdjb2x1bW4nLFxuXG4gICAgJyYgPiB1bC5tZWRpYS1saXN0Jzoge1xuICAgICAgLy8gVE9ETzogW1AyXSBXZSBzaG91bGQgbm90IHNldCBcImxpc3RTdHlsZVR5cGVcIiBpbiBzdHlsZVNldCwgdGhlIHVzZXIgaGF2ZSBubyBnb29kIHJlYXNvbnMgdG8gY2hhbmdlIHRoaXMuXG4gICAgICBsaXN0U3R5bGVUeXBlOiAnbm9uZScsXG4gICAgICBtYXJnaW46IHBhZGRpbmdSZWd1bGFyLFxuICAgICAgcGFkZGluZzogMFxuICAgIH1cbiAgfTtcbn1cbiJdLCJtYXBwaW5ncyI6Ijs7Ozs7OztBQUFBO0FBQ0E7QUFDQTtBQUVlLFNBQVNBLG1CQUFULE9BQTZEO0VBQUEsK0JBQTlCQyxjQUE4QjtFQUFBLElBQTlCQSxjQUE4QixvQ0FBYkMsU0FBYTtFQUMxRSxPQUFPO0lBQ0w7SUFDQUMsT0FBTyxFQUFFLE1BRko7SUFHTEMsYUFBYSxFQUFFLFFBSFY7SUFLTCxxQkFBcUI7TUFDbkI7TUFDQUMsYUFBYSxFQUFFLE1BRkk7TUFHbkJDLE1BQU0sRUFBRUwsY0FIVztNQUluQk0sT0FBTyxFQUFFO0lBSlU7RUFMaEIsQ0FBUDtBQVlEIn0=