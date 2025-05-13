"use strict";

var _interopRequireDefault = require("@babel/runtime/helpers/interopRequireDefault");

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.default = bunchUndos;

var _toConsumableArray2 = _interopRequireDefault(require("@babel/runtime/helpers/toConsumableArray"));

function bunchUndos(fns) {
  var called;
  return function () {
    if (!called) {
      called = true;
      (0, _toConsumableArray2.default)(fns).reverse().forEach(function (fn) {
        return fn === null || fn === void 0 ? void 0 : fn();
      });
    }
  };
}
//# sourceMappingURL=data:application/json;charset=utf-8;base64,eyJ2ZXJzaW9uIjozLCJuYW1lcyI6WyJidW5jaFVuZG9zIiwiZm5zIiwiY2FsbGVkIiwicmV2ZXJzZSIsImZvckVhY2giLCJmbiJdLCJzb3VyY2VSb290IjoiYnVuZGxlOi8vLyIsInNvdXJjZXMiOlsiLi4vLi4vLi4vc3JjL2FkYXB0aXZlQ2FyZHMvRE9NTWFuaXB1bGF0aW9uV2l0aFVuZG8vYnVuY2hVbmRvcy50c3giXSwic291cmNlc0NvbnRlbnQiOlsiaW1wb3J0IHR5cGUgeyBVbmRvRnVuY3Rpb24gfSBmcm9tICcuL3R5cGVzL1VuZG9GdW5jdGlvbic7XG5cbmV4cG9ydCBkZWZhdWx0IGZ1bmN0aW9uIGJ1bmNoVW5kb3MoZm5zOiBVbmRvRnVuY3Rpb25bXSk6IFVuZG9GdW5jdGlvbiB7XG4gIGxldCBjYWxsZWQ6IGJvb2xlYW47XG5cbiAgcmV0dXJuICgpID0+IHtcbiAgICBpZiAoIWNhbGxlZCkge1xuICAgICAgY2FsbGVkID0gdHJ1ZTtcbiAgICAgIFsuLi5mbnNdLnJldmVyc2UoKS5mb3JFYWNoKGZuID0+IGZuPy4oKSk7XG4gICAgfVxuICB9O1xufVxuIl0sIm1hcHBpbmdzIjoiOzs7Ozs7Ozs7OztBQUVlLFNBQVNBLFVBQVQsQ0FBb0JDLEdBQXBCLEVBQXVEO0VBQ3BFLElBQUlDLE1BQUo7RUFFQSxPQUFPLFlBQU07SUFDWCxJQUFJLENBQUNBLE1BQUwsRUFBYTtNQUNYQSxNQUFNLEdBQUcsSUFBVDtNQUNBLGlDQUFJRCxHQUFKLEVBQVNFLE9BQVQsR0FBbUJDLE9BQW5CLENBQTJCLFVBQUFDLEVBQUU7UUFBQSxPQUFJQSxFQUFKLGFBQUlBLEVBQUosdUJBQUlBLEVBQUUsRUFBTjtNQUFBLENBQTdCO0lBQ0Q7RUFDRixDQUxEO0FBTUQifQ==