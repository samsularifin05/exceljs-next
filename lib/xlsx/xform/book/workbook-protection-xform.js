'use strict';

const _ = require('../../../utils/under-dash');
const BaseXform = require('../base-xform');

function booleanToXml(model, value) {
  return model ? value : undefined;
}

function xmlToBoolean(value, equals) {
  return value === equals ? true : undefined;
}

class WorkbookProtectionXform extends BaseXform {
  get tag() {
    return 'workbookProtection';
  }

  render(xmlStream, model) {
    if (model) {
      const attributes = {
        lockStructure: booleanToXml(model.lockStructure, '1'),
        lockWindows: booleanToXml(model.lockWindows, '1'),
        lockRevision: booleanToXml(model.lockRevision, '1'),
        workbookAlgorithmName: model.algorithmName,
        workbookHashValue: model.hashValue,
        workbookSaltValue: model.saltValue,
        workbookSpinCount: model.spinCount,
      };
      if (_.some(attributes, value => value !== undefined)) {
        xmlStream.leafNode(this.tag, attributes);
      }
    }
  }

  parseOpen(node) {
    switch (node.name) {
      case this.tag:
        this.model = {};
        if (node.attributes.lockStructure) {
          this.model.lockStructure = xmlToBoolean(node.attributes.lockStructure, '1');
        }
        if (node.attributes.lockWindows) {
          this.model.lockWindows = xmlToBoolean(node.attributes.lockWindows, '1');
        }
        if (node.attributes.lockRevision) {
          this.model.lockRevision = xmlToBoolean(node.attributes.lockRevision, '1');
        }
        if (node.attributes.workbookAlgorithmName) {
          this.model.algorithmName = node.attributes.workbookAlgorithmName;
          this.model.hashValue = node.attributes.workbookHashValue;
          this.model.saltValue = node.attributes.workbookSaltValue;
          this.model.spinCount = parseInt(node.attributes.workbookSpinCount, 10);
        }
        return true;
      default:
        return false;
    }
  }

  parseText() {}

  parseClose() {
    return false;
  }
}

module.exports = WorkbookProtectionXform;
