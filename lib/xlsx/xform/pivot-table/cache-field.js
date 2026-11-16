const utils = require('../../../utils/utils');

class CacheField {
  constructor({name, sharedItems}) {
    this.name = name;
    this.sharedItems = sharedItems;
  }

  render() {
    // PivotCache Field: http://www.datypic.com/sc/ooxml/e-ssml_cacheField-1.html
    // Shared Items: http://www.datypic.com/sc/ooxml/e-ssml_sharedItems-1.html
    const escapedName = utils.xmlEncode(String(this.name != null ? this.name : ''));

    // integer types
    if (this.sharedItems === null) {
      // TK(2023-07-18): left out attributes... minValue="5" maxValue="45"
      return `<cacheField name="${escapedName}" numFmtId="0">
      <sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" />
    </cacheField>`;
    }

    // string types
    return `<cacheField name="${escapedName}" numFmtId="0">
      <sharedItems count="${this.sharedItems.length}">
        ${this.sharedItems
          .map(item =>
            item === null || item === undefined
              ? '<m />'
              : `<s v="${utils.xmlEncode(String(item))}" />`
          )
          .join('')}
      </sharedItems>
    </cacheField>`;
  }
}

module.exports = CacheField;
