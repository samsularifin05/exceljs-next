const BaseXform = require('../base-xform');

class PageBreaksXform extends BaseXform {
  get tag() {
    return 'brk';
  }

  render(xmlStream, model) {
    xmlStream.leafNode('brk', model);
  }

  parseOpen(node) {
    if (node.name === 'brk') {
      const model = {
        id: parseInt(node.attributes.id, 10),
        max: parseInt(node.attributes.max, 10),
        man: parseInt(node.attributes.man, 10),
      };
      if (node.attributes.min !== undefined) {
        model.min = parseInt(node.attributes.min, 10);
      }
      this.model = model;
      return true;
    }
    return false;
  }

  parseText() {}

  parseClose() {
    return false;
  }
}

module.exports = PageBreaksXform;
