"use strict";

var _interopRequireDefault = require("@babel/runtime/helpers/interopRequireDefault");

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.default = useUniqueId;

var _react = require("react");

var _mathRandom = _interopRequireDefault(require("math-random"));

/* eslint no-magic-numbers: ["error", { "ignore": [2, 5, 36] }] */
function useUniqueId(prefix) {
  var id = (0, _react.useMemo)(function () {
    return (0, _mathRandom.default)().toString(36).substr(2, 5);
  }, []);
  prefix = prefix ? "".concat(prefix, "--") : '';
  return "".concat(prefix).concat(id);
}
//# sourceMappingURL=data:application/json;charset=utf-8;base64,eyJ2ZXJzaW9uIjozLCJuYW1lcyI6WyJ1c2VVbmlxdWVJZCIsInByZWZpeCIsImlkIiwidXNlTWVtbyIsInJhbmRvbSIsInRvU3RyaW5nIiwic3Vic3RyIl0sInNvdXJjZVJvb3QiOiJidW5kbGU6Ly8vIiwic291cmNlcyI6WyIuLi8uLi8uLi8uLi9zcmMvYWRhcHRpdmVDYXJkcy9ob29rcy9pbnRlcm5hbC91c2VVbmlxdWVJZC50cyJdLCJzb3VyY2VzQ29udGVudCI6WyIvKiBlc2xpbnQgbm8tbWFnaWMtbnVtYmVyczogW1wiZXJyb3JcIiwgeyBcImlnbm9yZVwiOiBbMiwgNSwgMzZdIH1dICovXG5cbmltcG9ydCB7IHVzZU1lbW8gfSBmcm9tICdyZWFjdCc7XG5pbXBvcnQgcmFuZG9tIGZyb20gJ21hdGgtcmFuZG9tJztcblxuZXhwb3J0IGRlZmF1bHQgZnVuY3Rpb24gdXNlVW5pcXVlSWQocHJlZml4Pzogc3RyaW5nKTogc3RyaW5nIHtcbiAgY29uc3QgaWQgPSB1c2VNZW1vKCgpID0+IHJhbmRvbSgpLnRvU3RyaW5nKDM2KS5zdWJzdHIoMiwgNSksIFtdKTtcblxuICBwcmVmaXggPSBwcmVmaXggPyBgJHtwcmVmaXh9LS1gIDogJyc7XG5cbiAgcmV0dXJuIGAke3ByZWZpeH0ke2lkfWA7XG59XG4iXSwibWFwcGluZ3MiOiI7Ozs7Ozs7OztBQUVBOztBQUNBOztBQUhBO0FBS2UsU0FBU0EsV0FBVCxDQUFxQkMsTUFBckIsRUFBOEM7RUFDM0QsSUFBTUMsRUFBRSxHQUFHLElBQUFDLGNBQUEsRUFBUTtJQUFBLE9BQU0sSUFBQUMsbUJBQUEsSUFBU0MsUUFBVCxDQUFrQixFQUFsQixFQUFzQkMsTUFBdEIsQ0FBNkIsQ0FBN0IsRUFBZ0MsQ0FBaEMsQ0FBTjtFQUFBLENBQVIsRUFBa0QsRUFBbEQsQ0FBWDtFQUVBTCxNQUFNLEdBQUdBLE1BQU0sYUFBTUEsTUFBTixVQUFtQixFQUFsQztFQUVBLGlCQUFVQSxNQUFWLFNBQW1CQyxFQUFuQjtBQUNEIn0=