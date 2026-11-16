const BaseCellAnchorXform = require('./base-cell-anchor-xform');
const StaticXform = require('../static-xform');

const CellPositionXform = require('./cell-position-xform');
const PicXform = require('./pic-xform');

class TwoCellAnchorXform extends BaseCellAnchorXform {
  constructor() {
    super();

    this.map = {
      'xdr:from': new CellPositionXform({tag: 'xdr:from'}),
      'xdr:to': new CellPositionXform({tag: 'xdr:to'}),
      'xdr:pic': new PicXform(),
      'xdr:clientData': new StaticXform({tag: 'xdr:clientData'}),
    };
  }

  get tag() {
    return 'xdr:twoCellAnchor';
  }

  prepare(model, options) {
    if (model.picture) {
      this.map['xdr:pic'].prepare(model.picture, options);
    }
  }

  render(xmlStream, model) {
    xmlStream.openNode(this.tag, {editAs: model.range.editAs || 'oneCell'});

    this.map['xdr:from'].render(xmlStream, model.range.tl);
    this.map['xdr:to'].render(xmlStream, model.range.br);
    if (model.chart) {
      xmlStream.openNode('xdr:graphicFrame', {macro: ''});
      xmlStream.openNode('xdr:nvGraphicFramePr');
      xmlStream.leafNode('xdr:cNvPr', {
        id: String(model.chart.id || 2),
        name: model.chart.name || `Chart ${model.chart.id || 1}`,
      });
      xmlStream.leafNode('xdr:cNvGraphicFramePr');
      xmlStream.closeNode(); // xdr:nvGraphicFramePr
      xmlStream.openNode('xdr:xfrm');
      xmlStream.leafNode('a:off', {x: '0', y: '0'});
      xmlStream.leafNode('a:ext', {cx: '0', cy: '0'});
      xmlStream.closeNode(); // xdr:xfrm
      xmlStream.openNode('a:graphic');
      xmlStream.openNode('a:graphicData', {
        uri: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
      });
      xmlStream.leafNode('c:chart', {
        'xmlns:c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
        'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'r:id': model.chart.rId,
      });
      xmlStream.closeNode(); // a:graphicData
      xmlStream.closeNode(); // a:graphic
      xmlStream.closeNode(); // xdr:graphicFrame
    } else {
      this.map['xdr:pic'].render(xmlStream, model.picture);
    }
    this.map['xdr:clientData'].render(xmlStream, {});

    xmlStream.closeNode();
  }

  parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        this.model.range.tl = this.map['xdr:from'].model;
        this.model.range.br = this.map['xdr:to'].model;
        this.model.picture = this.map['xdr:pic'].model;
        return false;
      default:
        // could be some unrecognised tags
        return true;
    }
  }

  reconcile(model, options) {
    model.medium = this.reconcilePicture(model.picture, options);
  }
}

module.exports = TwoCellAnchorXform;
