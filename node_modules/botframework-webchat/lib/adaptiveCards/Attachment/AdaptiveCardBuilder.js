"use strict";

var _interopRequireDefault = require("@babel/runtime/helpers/interopRequireDefault");

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.default = void 0;

var _classCallCheck2 = _interopRequireDefault(require("@babel/runtime/helpers/classCallCheck"));

var _createClass2 = _interopRequireDefault(require("@babel/runtime/helpers/createClass"));

var _defineProperty2 = _interopRequireDefault(require("@babel/runtime/helpers/defineProperty"));

var _adaptivecards = require("adaptivecards");

var _botframeworkWebchatCore = require("botframework-webchat-core");

function addCardAction(cardAction, includesOAuthButtons) {
  var type = cardAction.type;
  var action;

  if (type === 'imBack' || type === 'messageBack' || type === 'postBack' || type === 'signin' && includesOAuthButtons) {
    action = new _adaptivecards.SubmitAction();
    action.data = {
      __isBotFrameworkCardAction: true,
      cardAction: cardAction
    };
    action.title = cardAction.title;
  } else {
    action = new _adaptivecards.OpenUrlAction();
    action.title = cardAction.title;
    action.url = cardAction.type === 'call' ? "tel:".concat(cardAction.value) : cardAction.value;
  }

  return action;
}

var AdaptiveCardBuilder = /*#__PURE__*/function () {
  function AdaptiveCardBuilder(adaptiveCards, styleOptions) {
    var direction = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : 'ltr';
    (0, _classCallCheck2.default)(this, AdaptiveCardBuilder);
    (0, _defineProperty2.default)(this, "card", void 0);
    (0, _defineProperty2.default)(this, "container", void 0);
    (0, _defineProperty2.default)(this, "styleOptions", void 0);
    this.card = new adaptiveCards.AdaptiveCard();
    this.container = new _adaptivecards.Container();
    this.container.rtl = direction === 'rtl';
    this.styleOptions = styleOptions;
    this.card.addItem(this.container);
  }

  (0, _createClass2.default)(AdaptiveCardBuilder, [{
    key: "addColumnSet",
    value: function addColumnSet(sizes) {
      var container = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : this.container;
      var selectAction = arguments.length > 2 ? arguments[2] : undefined;
      var columnSet = new _adaptivecards.ColumnSet();
      columnSet.selectAction = selectAction && addCardAction(selectAction);
      container.addItem(columnSet);
      return sizes.map(function (size) {
        var column = new _adaptivecards.Column();
        column.width = new _adaptivecards.SizeAndUnit(size, _adaptivecards.SizeUnit.Weight);
        columnSet.addColumn(column);
        return column;
      });
    }
  }, {
    key: "addItems",
    value: function addItems(cardElements) {
      var container = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : this.container;
      cardElements.forEach(function (cardElement) {
        return container.addItem(cardElement);
      });
    }
  }, {
    key: "addTextBlock",
    value: function addTextBlock(text, template) {
      var container = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : this.container;

      if (typeof text !== 'undefined') {
        var textblock = new _adaptivecards.TextBlock();

        for (var prop in template) {
          if (!(0, _botframeworkWebchatCore.isForbiddenPropertyName)(prop)) {
            // Mitigated through denylisting.
            // eslint-disable-next-line security/detect-object-injection
            textblock[prop] = template[prop];
          }
        }

        textblock.text = text;
        container.addItem(textblock);
      }
    }
  }, {
    key: "addButtons",
    value: function addButtons(cardActions, includesOAuthButtons) {
      var _this = this;

      cardActions && cardActions.forEach(function (cardAction) {
        _this.card.addAction(addCardAction(cardAction, includesOAuthButtons));
      });
    }
  }, {
    key: "addCommonHeaders",
    value: function addCommonHeaders(content) {
      var richCardWrapTitle = this.styleOptions.richCardWrapTitle;
      this.addTextBlock(content.title, {
        color: _adaptivecards.TextColor.Default,
        size: _adaptivecards.TextSize.Medium,
        style: 'heading',
        weight: _adaptivecards.TextWeight.Bolder,
        wrap: richCardWrapTitle
      });
      this.addTextBlock(content.subtitle, {
        color: _adaptivecards.TextColor.Default,
        isSubtle: true,
        wrap: richCardWrapTitle
      });
      this.addTextBlock(content.text, {
        color: _adaptivecards.TextColor.Default,
        wrap: true
      });
    }
  }, {
    key: "addCommon",
    value: function addCommon(content) {
      this.addCommonHeaders(content);
      this.addButtons(content.buttons);
    }
  }, {
    key: "addImage",
    value: function addImage(url, container, selectAction, altText) {
      container = container || this.container;
      var image = new _adaptivecards.Image();
      image.altText = altText;
      image.url = url;
      image.selectAction = selectAction && addCardAction(selectAction);
      image.size = _adaptivecards.Size.Stretch;
      container.addItem(image);
    }
  }]);
  return AdaptiveCardBuilder;
}();

exports.default = AdaptiveCardBuilder;
//# sourceMappingURL=data:application/json;charset=utf-8;base64,eyJ2ZXJzaW9uIjozLCJuYW1lcyI6WyJhZGRDYXJkQWN0aW9uIiwiY2FyZEFjdGlvbiIsImluY2x1ZGVzT0F1dGhCdXR0b25zIiwidHlwZSIsImFjdGlvbiIsIlN1Ym1pdEFjdGlvbiIsImRhdGEiLCJfX2lzQm90RnJhbWV3b3JrQ2FyZEFjdGlvbiIsInRpdGxlIiwiT3BlblVybEFjdGlvbiIsInVybCIsInZhbHVlIiwiQWRhcHRpdmVDYXJkQnVpbGRlciIsImFkYXB0aXZlQ2FyZHMiLCJzdHlsZU9wdGlvbnMiLCJkaXJlY3Rpb24iLCJjYXJkIiwiQWRhcHRpdmVDYXJkIiwiY29udGFpbmVyIiwiQ29udGFpbmVyIiwicnRsIiwiYWRkSXRlbSIsInNpemVzIiwic2VsZWN0QWN0aW9uIiwiY29sdW1uU2V0IiwiQ29sdW1uU2V0IiwibWFwIiwic2l6ZSIsImNvbHVtbiIsIkNvbHVtbiIsIndpZHRoIiwiU2l6ZUFuZFVuaXQiLCJTaXplVW5pdCIsIldlaWdodCIsImFkZENvbHVtbiIsImNhcmRFbGVtZW50cyIsImZvckVhY2giLCJjYXJkRWxlbWVudCIsInRleHQiLCJ0ZW1wbGF0ZSIsInRleHRibG9jayIsIlRleHRCbG9jayIsInByb3AiLCJpc0ZvcmJpZGRlblByb3BlcnR5TmFtZSIsImNhcmRBY3Rpb25zIiwiYWRkQWN0aW9uIiwiY29udGVudCIsInJpY2hDYXJkV3JhcFRpdGxlIiwiYWRkVGV4dEJsb2NrIiwiY29sb3IiLCJUZXh0Q29sb3IiLCJEZWZhdWx0IiwiVGV4dFNpemUiLCJNZWRpdW0iLCJzdHlsZSIsIndlaWdodCIsIlRleHRXZWlnaHQiLCJCb2xkZXIiLCJ3cmFwIiwic3VidGl0bGUiLCJpc1N1YnRsZSIsImFkZENvbW1vbkhlYWRlcnMiLCJhZGRCdXR0b25zIiwiYnV0dG9ucyIsImFsdFRleHQiLCJpbWFnZSIsIkltYWdlIiwiU2l6ZSIsIlN0cmV0Y2giXSwic291cmNlUm9vdCI6ImJ1bmRsZTovLy8iLCJzb3VyY2VzIjpbIi4uLy4uLy4uL3NyYy9hZGFwdGl2ZUNhcmRzL0F0dGFjaG1lbnQvQWRhcHRpdmVDYXJkQnVpbGRlci50cyJdLCJzb3VyY2VzQ29udGVudCI6WyJpbXBvcnQge1xuICBBZGFwdGl2ZUNhcmQsXG4gIENhcmRFbGVtZW50LFxuICBDb2x1bW4sXG4gIENvbHVtblNldCxcbiAgQ29udGFpbmVyLFxuICBJbWFnZSxcbiAgT3BlblVybEFjdGlvbixcbiAgU2l6ZSxcbiAgU2l6ZUFuZFVuaXQsXG4gIFNpemVVbml0LFxuICBTdWJtaXRBY3Rpb24sXG4gIFRleHRCbG9jayxcbiAgVGV4dENvbG9yLFxuICBUZXh0U2l6ZSxcbiAgVGV4dFdlaWdodFxufSBmcm9tICdhZGFwdGl2ZWNhcmRzJztcbmltcG9ydCB7IGlzRm9yYmlkZGVuUHJvcGVydHlOYW1lIH0gZnJvbSAnYm90ZnJhbWV3b3JrLXdlYmNoYXQtY29yZSc7XG5pbXBvcnQgdHlwZSB7IERpcmVjdExpbmVDYXJkQWN0aW9uIH0gZnJvbSAnYm90ZnJhbWV3b3JrLXdlYmNoYXQtY29yZSc7XG5cbmltcG9ydCBBZGFwdGl2ZUNhcmRzUGFja2FnZSBmcm9tICcuLi8uLi90eXBlcy9BZGFwdGl2ZUNhcmRzUGFja2FnZSc7XG5pbXBvcnQgQWRhcHRpdmVDYXJkc1N0eWxlT3B0aW9ucyBmcm9tICcuLi9BZGFwdGl2ZUNhcmRzU3R5bGVPcHRpb25zJztcblxuZXhwb3J0IGludGVyZmFjZSBCb3RGcmFtZXdvcmtDYXJkQWN0aW9uIHtcbiAgX19pc0JvdEZyYW1ld29ya0NhcmRBY3Rpb246IHRydWU7XG4gIGNhcmRBY3Rpb246IERpcmVjdExpbmVDYXJkQWN0aW9uO1xufVxuXG5mdW5jdGlvbiBhZGRDYXJkQWN0aW9uKGNhcmRBY3Rpb246IERpcmVjdExpbmVDYXJkQWN0aW9uLCBpbmNsdWRlc09BdXRoQnV0dG9ucz86IGJvb2xlYW4pIHtcbiAgY29uc3QgeyB0eXBlIH0gPSBjYXJkQWN0aW9uO1xuICBsZXQgYWN0aW9uO1xuXG4gIGlmIChcbiAgICB0eXBlID09PSAnaW1CYWNrJyB8fFxuICAgIHR5cGUgPT09ICdtZXNzYWdlQmFjaycgfHxcbiAgICB0eXBlID09PSAncG9zdEJhY2snIHx8XG4gICAgKHR5cGUgPT09ICdzaWduaW4nICYmIGluY2x1ZGVzT0F1dGhCdXR0b25zKVxuICApIHtcbiAgICBhY3Rpb24gPSBuZXcgU3VibWl0QWN0aW9uKCk7XG5cbiAgICBhY3Rpb24uZGF0YSA9IHtcbiAgICAgIF9faXNCb3RGcmFtZXdvcmtDYXJkQWN0aW9uOiB0cnVlLFxuICAgICAgY2FyZEFjdGlvblxuICAgIH07XG5cbiAgICBhY3Rpb24udGl0bGUgPSAoY2FyZEFjdGlvbiBhcyB7IHRpdGxlOiBzdHJpbmcgfSkudGl0bGU7XG4gIH0gZWxzZSB7XG4gICAgYWN0aW9uID0gbmV3IE9wZW5VcmxBY3Rpb24oKTtcblxuICAgIGFjdGlvbi50aXRsZSA9IChjYXJkQWN0aW9uIGFzIHsgdGl0bGU6IHN0cmluZyB9KS50aXRsZTtcbiAgICBhY3Rpb24udXJsID0gY2FyZEFjdGlvbi50eXBlID09PSAnY2FsbCcgPyBgdGVsOiR7Y2FyZEFjdGlvbi52YWx1ZX1gIDogY2FyZEFjdGlvbi52YWx1ZTtcbiAgfVxuXG4gIHJldHVybiBhY3Rpb247XG59XG5cbmV4cG9ydCBkZWZhdWx0IGNsYXNzIEFkYXB0aXZlQ2FyZEJ1aWxkZXIge1xuICBjYXJkOiBBZGFwdGl2ZUNhcmQ7XG4gIGNvbnRhaW5lcjogQ29udGFpbmVyO1xuICBzdHlsZU9wdGlvbnM6IEFkYXB0aXZlQ2FyZHNTdHlsZU9wdGlvbnM7XG5cbiAgY29uc3RydWN0b3IoXG4gICAgYWRhcHRpdmVDYXJkczogQWRhcHRpdmVDYXJkc1BhY2thZ2UsXG4gICAgc3R5bGVPcHRpb25zOiBBZGFwdGl2ZUNhcmRzU3R5bGVPcHRpb25zLFxuICAgIGRpcmVjdGlvbjogJ2x0cicgfCAncnRsJyB8ICdhdXRvJyA9ICdsdHInXG4gICkge1xuICAgIHRoaXMuY2FyZCA9IG5ldyBhZGFwdGl2ZUNhcmRzLkFkYXB0aXZlQ2FyZCgpO1xuICAgIHRoaXMuY29udGFpbmVyID0gbmV3IENvbnRhaW5lcigpO1xuICAgIHRoaXMuY29udGFpbmVyLnJ0bCA9IGRpcmVjdGlvbiA9PT0gJ3J0bCc7XG4gICAgdGhpcy5zdHlsZU9wdGlvbnMgPSBzdHlsZU9wdGlvbnM7XG5cbiAgICB0aGlzLmNhcmQuYWRkSXRlbSh0aGlzLmNvbnRhaW5lcik7XG4gIH1cblxuICBhZGRDb2x1bW5TZXQoc2l6ZXM6IG51bWJlcltdLCBjb250YWluZXI6IENvbnRhaW5lciA9IHRoaXMuY29udGFpbmVyLCBzZWxlY3RBY3Rpb24/OiBEaXJlY3RMaW5lQ2FyZEFjdGlvbikge1xuICAgIGNvbnN0IGNvbHVtblNldCA9IG5ldyBDb2x1bW5TZXQoKTtcblxuICAgIGNvbHVtblNldC5zZWxlY3RBY3Rpb24gPSBzZWxlY3RBY3Rpb24gJiYgYWRkQ2FyZEFjdGlvbihzZWxlY3RBY3Rpb24pO1xuICAgIGNvbnRhaW5lci5hZGRJdGVtKGNvbHVtblNldCk7XG5cbiAgICByZXR1cm4gc2l6ZXMubWFwKHNpemUgPT4ge1xuICAgICAgY29uc3QgY29sdW1uID0gbmV3IENvbHVtbigpO1xuXG4gICAgICBjb2x1bW4ud2lkdGggPSBuZXcgU2l6ZUFuZFVuaXQoc2l6ZSwgU2l6ZVVuaXQuV2VpZ2h0KTtcblxuICAgICAgY29sdW1uU2V0LmFkZENvbHVtbihjb2x1bW4pO1xuXG4gICAgICByZXR1cm4gY29sdW1uO1xuICAgIH0pO1xuICB9XG5cbiAgYWRkSXRlbXMoY2FyZEVsZW1lbnRzOiBDYXJkRWxlbWVudFtdLCBjb250YWluZXI6IENvbnRhaW5lciA9IHRoaXMuY29udGFpbmVyKSB7XG4gICAgY2FyZEVsZW1lbnRzLmZvckVhY2goY2FyZEVsZW1lbnQgPT4gY29udGFpbmVyLmFkZEl0ZW0oY2FyZEVsZW1lbnQpKTtcbiAgfVxuXG4gIGFkZFRleHRCbG9jayh0ZXh0OiBzdHJpbmcsIHRlbXBsYXRlOiBQYXJ0aWFsPFRleHRCbG9jaz4sIGNvbnRhaW5lcjogQ29udGFpbmVyID0gdGhpcy5jb250YWluZXIpIHtcbiAgICBpZiAodHlwZW9mIHRleHQgIT09ICd1bmRlZmluZWQnKSB7XG4gICAgICBjb25zdCB0ZXh0YmxvY2sgPSBuZXcgVGV4dEJsb2NrKCk7XG5cbiAgICAgIGZvciAoY29uc3QgcHJvcCBpbiB0ZW1wbGF0ZSkge1xuICAgICAgICBpZiAoIWlzRm9yYmlkZGVuUHJvcGVydHlOYW1lKHByb3ApKSB7XG4gICAgICAgICAgLy8gTWl0aWdhdGVkIHRocm91Z2ggZGVueWxpc3RpbmcuXG4gICAgICAgICAgLy8gZXNsaW50LWRpc2FibGUtbmV4dC1saW5lIHNlY3VyaXR5L2RldGVjdC1vYmplY3QtaW5qZWN0aW9uXG4gICAgICAgICAgdGV4dGJsb2NrW3Byb3BdID0gdGVtcGxhdGVbcHJvcF07XG4gICAgICAgIH1cbiAgICAgIH1cblxuICAgICAgdGV4dGJsb2NrLnRleHQgPSB0ZXh0O1xuXG4gICAgICBjb250YWluZXIuYWRkSXRlbSh0ZXh0YmxvY2spO1xuICAgIH1cbiAgfVxuXG4gIGFkZEJ1dHRvbnMoY2FyZEFjdGlvbnM6IERpcmVjdExpbmVDYXJkQWN0aW9uW10sIGluY2x1ZGVzT0F1dGhCdXR0b25zPzogYm9vbGVhbikge1xuICAgIGNhcmRBY3Rpb25zICYmXG4gICAgICBjYXJkQWN0aW9ucy5mb3JFYWNoKGNhcmRBY3Rpb24gPT4ge1xuICAgICAgICB0aGlzLmNhcmQuYWRkQWN0aW9uKGFkZENhcmRBY3Rpb24oY2FyZEFjdGlvbiwgaW5jbHVkZXNPQXV0aEJ1dHRvbnMpKTtcbiAgICAgIH0pO1xuICB9XG5cbiAgYWRkQ29tbW9uSGVhZGVycyhjb250ZW50OiBJQ29tbW9uQ29udGVudCkge1xuICAgIGNvbnN0IHsgcmljaENhcmRXcmFwVGl0bGUgfSA9IHRoaXMuc3R5bGVPcHRpb25zO1xuICAgIHRoaXMuYWRkVGV4dEJsb2NrKGNvbnRlbnQudGl0bGUsIHtcbiAgICAgIGNvbG9yOiBUZXh0Q29sb3IuRGVmYXVsdCxcbiAgICAgIHNpemU6IFRleHRTaXplLk1lZGl1bSxcbiAgICAgIHN0eWxlOiAnaGVhZGluZycsXG4gICAgICB3ZWlnaHQ6IFRleHRXZWlnaHQuQm9sZGVyLFxuICAgICAgd3JhcDogcmljaENhcmRXcmFwVGl0bGVcbiAgICB9KTtcbiAgICB0aGlzLmFkZFRleHRCbG9jayhjb250ZW50LnN1YnRpdGxlLCB7IGNvbG9yOiBUZXh0Q29sb3IuRGVmYXVsdCwgaXNTdWJ0bGU6IHRydWUsIHdyYXA6IHJpY2hDYXJkV3JhcFRpdGxlIH0pO1xuICAgIHRoaXMuYWRkVGV4dEJsb2NrKGNvbnRlbnQudGV4dCwgeyBjb2xvcjogVGV4dENvbG9yLkRlZmF1bHQsIHdyYXA6IHRydWUgfSk7XG4gIH1cblxuICBhZGRDb21tb24oY29udGVudDogSUNvbW1vbkNvbnRlbnQpIHtcbiAgICB0aGlzLmFkZENvbW1vbkhlYWRlcnMoY29udGVudCk7XG4gICAgdGhpcy5hZGRCdXR0b25zKGNvbnRlbnQuYnV0dG9ucyk7XG4gIH1cblxuICBhZGRJbWFnZSh1cmw6IHN0cmluZywgY29udGFpbmVyPzogQ29udGFpbmVyLCBzZWxlY3RBY3Rpb24/OiBEaXJlY3RMaW5lQ2FyZEFjdGlvbiwgYWx0VGV4dD86IHN0cmluZykge1xuICAgIGNvbnRhaW5lciA9IGNvbnRhaW5lciB8fCB0aGlzLmNvbnRhaW5lcjtcblxuICAgIGNvbnN0IGltYWdlID0gbmV3IEltYWdlKCk7XG5cbiAgICBpbWFnZS5hbHRUZXh0ID0gYWx0VGV4dDtcbiAgICBpbWFnZS51cmwgPSB1cmw7XG4gICAgaW1hZ2Uuc2VsZWN0QWN0aW9uID0gc2VsZWN0QWN0aW9uICYmIGFkZENhcmRBY3Rpb24oc2VsZWN0QWN0aW9uKTtcbiAgICBpbWFnZS5zaXplID0gU2l6ZS5TdHJldGNoO1xuXG4gICAgY29udGFpbmVyLmFkZEl0ZW0oaW1hZ2UpO1xuICB9XG59XG5cbmV4cG9ydCBpbnRlcmZhY2UgSUNvbW1vbkNvbnRlbnQge1xuICBidXR0b25zPzogRGlyZWN0TGluZUNhcmRBY3Rpb25bXTtcbiAgc3VidGl0bGU/OiBzdHJpbmc7XG4gIHRleHQ/OiBzdHJpbmc7XG4gIHRpdGxlPzogc3RyaW5nO1xufVxuIl0s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