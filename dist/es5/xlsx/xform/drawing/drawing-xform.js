/**
 * Copyright (c) 2016-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var _typeof = typeof Symbol === "function" && typeof Symbol.iterator === "symbol" ? function (obj) { return typeof obj; } : function (obj) { return obj && typeof Symbol === "function" && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj; };

var utils = require('../../../utils/utils');
var XmlStream = require('../../../utils/xml-stream');

var BaseXform = require('../base-xform');
var TwoCellAnchorXform = require('./two-cell-anchor-xform');
var OneCellAnchorXform = require('./one-cell-anchor-xform');

var WorkSheetXform = module.exports = function () {
  this.map = {
    'xdr:twoCellAnchor': new TwoCellAnchorXform(),
    'xdr:oneCellAnchor': new OneCellAnchorXform()
  };
};

function useOneCellAnchor(model) {
  return _typeof(model.range) === 'object' && model.range.ext;
}

function reconcileOneCellAnchor(model) {
  return !!model.ext;
}

utils.inherits(WorkSheetXform, BaseXform, {
  DRAWING_ATTRIBUTES: {
    'xmlns:xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
  }
}, {
  get tag() {
    return 'xdr:wsDr';
  },

  prepare: function prepare(model) {
    var twoCellAnchorXform = this.map['xdr:twoCellAnchor'];
    model.anchors.forEach(function (item, index) {
      if (!useOneCellAnchor(item)) {
        twoCellAnchorXform.prepare(item, { index: index });
      }
    });

    var oneCellAnchorXform = this.map['xdr:oneCellAnchor'];
    model.anchors.forEach(function (item, index) {
      if (useOneCellAnchor(item)) {
        oneCellAnchorXform.prepare(item, { index: index });
      }
    });
  },

  render: function render(xmlStream, model) {
    xmlStream.openXml(XmlStream.StdDocAttributes);
    xmlStream.openNode(this.tag, WorkSheetXform.DRAWING_ATTRIBUTES);

    var twoCellAnchorXform = this.map['xdr:twoCellAnchor'];
    model.anchors.forEach(function (item) {
      if (!useOneCellAnchor(item)) {
        twoCellAnchorXform.render(xmlStream, item);
      }
    });

    var oneCellAnchorXform = this.map['xdr:oneCellAnchor'];
    model.anchors.forEach(function (item) {
      if (useOneCellAnchor(item)) {
        oneCellAnchorXform.render(xmlStream, item);
      }
    });

    xmlStream.closeNode();
  },

  parseOpen: function parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.reset();
        this.model = {
          anchors: []
        };
        break;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        break;
    }
    return true;
  },

  parseText: function parseText(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  },

  parseClose: function parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.model.anchors.push(this.parser.model);
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        return false;
      default:
        // could be some unrecognised tags
        return true;
    }
  },

  reconcile: function reconcile(model, options) {
    var _this = this;

    model.anchors.forEach(function (anchor) {
      if (reconcileOneCellAnchor(anchor)) {
        _this.map['xdr:oneCellAnchor'].reconcile(anchor, options);
      } else {
        _this.map['xdr:twoCellAnchor'].reconcile(anchor, options);
      }
    });
  }
});
//# sourceMappingURL=drawing-xform.js.map
