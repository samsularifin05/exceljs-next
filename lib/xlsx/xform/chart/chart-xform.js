'use strict';

const XmlStream = require('../../../utils/xml-stream');
const BaseXform = require('../base-xform');

const DEFAULT_COLORS = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];

function qualifyRange(rangeStr, sheetName) {
  if (!rangeStr) {
    return '';
  }
  const str = String(rangeStr).trim();
  if (str.includes('!')) {
    return str;
  }
  const escapedSheet = String(sheetName || 'Sheet1').replace(/'/g, '\'\'');
  return `'${escapedSheet}'!${str}`;
}

function normalizeColor(color) {
  if (!color) {
    return null;
  }
  return String(color).replace(/^#/, '').toUpperCase().slice(-6);
}

class ChartXform extends BaseXform {
  get tag() {
    return 'c:chartSpace';
  }

  render(xmlStream, model) {
    const type = (model.type || 'column').toLowerCase();
    const grouping = model.grouping || (type === 'line' || type === 'area' ? 'standard' : 'clustered');
    const legendPos = (model.legend && model.legend.position) || 'b';
    const valAxisNumFmt = model.valAxis && model.valAxis.numFmt;
    const valAxisMajorUnit = model.valAxis && model.valAxis.majorUnit;
    const seriesList = Array.isArray(model.series) ? model.series : [];
    const sheetName = model.sheetName || 'Sheet1';
    const catAxisId = 100001;
    const valAxisId = 100002;
    const isPieLike = type === 'pie' || type === 'doughnut';

    xmlStream.openXml(XmlStream.StdDocAttributes);
    xmlStream.openNode('c:chartSpace', ChartXform.CHART_SPACE_ATTRIBUTES);
    xmlStream.leafNode('c:date1904', {val: '0'});
    xmlStream.leafNode('c:lang', {val: 'en-US'});
    xmlStream.leafNode('c:roundedCorners', {val: '0'});

    xmlStream.openNode('c:chart');
    if (model.title) {
      xmlStream.openNode('c:title');
      xmlStream.openNode('c:tx');
      xmlStream.openNode('c:rich');
      xmlStream.leafNode('a:bodyPr');
      xmlStream.leafNode('a:lstStyle');
      xmlStream.openNode('a:p');
      xmlStream.openNode('a:r');
      xmlStream.leafNode('a:rPr', {lang: 'en-US', sz: '1400', b: '1'});
      xmlStream.leafNode('a:t', null, String(model.title));
      xmlStream.closeNode(); // a:r
      xmlStream.closeNode(); // a:p
      xmlStream.closeNode(); // c:rich
      xmlStream.closeNode(); // c:tx
      xmlStream.leafNode('c:overlay', {val: '0'});
      xmlStream.closeNode(); // c:title
    } else {
      xmlStream.leafNode('c:autoTitleDeleted', {val: '1'});
    }

    xmlStream.openNode('c:plotArea');
    xmlStream.leafNode('c:layout');

    // Render chart-type node
    if (type === 'bar' || type === 'column') {
      xmlStream.openNode('c:barChart');
      xmlStream.leafNode('c:barDir', {val: type === 'bar' ? 'bar' : 'col'});
      xmlStream.leafNode('c:grouping', {val: grouping});
      xmlStream.leafNode('c:varyColors', {val: '0'});
    } else if (type === 'line') {
      xmlStream.openNode('c:lineChart');
      xmlStream.leafNode('c:grouping', {val: grouping});
      xmlStream.leafNode('c:varyColors', {val: '0'});
    } else if (type === 'area') {
      xmlStream.openNode('c:areaChart');
      xmlStream.leafNode('c:grouping', {val: grouping});
      xmlStream.leafNode('c:varyColors', {val: '0'});
    } else if (type === 'doughnut') {
      xmlStream.openNode('c:doughnutChart');
      xmlStream.leafNode('c:varyColors', {val: '1'});
    } else {
      xmlStream.openNode('c:pieChart');
      xmlStream.leafNode('c:varyColors', {val: '1'});
    }

    seriesList.forEach((ser, idx) => {
      xmlStream.openNode('c:ser');
      xmlStream.leafNode('c:idx', {val: String(idx)});
      xmlStream.leafNode('c:order', {val: String(idx)});

      if (ser.name) {
        xmlStream.openNode('c:tx');
        xmlStream.leafNode('c:v', null, String(ser.name));
        xmlStream.closeNode();
      }

      const color = normalizeColor(ser.color) || DEFAULT_COLORS[idx % DEFAULT_COLORS.length];
      if (!isPieLike) {
        xmlStream.openNode('c:spPr');
        if (type === 'line') {
          xmlStream.openNode('a:ln', {w: '28575'});
          xmlStream.openNode('a:solidFill');
          xmlStream.leafNode('a:srgbClr', {val: color});
          xmlStream.closeNode();
          xmlStream.closeNode();
        } else {
          xmlStream.openNode('a:solidFill');
          xmlStream.leafNode('a:srgbClr', {val: color});
          xmlStream.closeNode();
        }
        xmlStream.closeNode();
      }

      if (ser.categories) {
        xmlStream.openNode('c:cat');
        xmlStream.openNode('c:strRef');
        xmlStream.leafNode('c:f', null, qualifyRange(ser.categories, sheetName));
        xmlStream.closeNode();
        xmlStream.closeNode();
      }

      if (ser.values) {
        xmlStream.openNode('c:val');
        xmlStream.openNode('c:numRef');
        xmlStream.leafNode('c:f', null, qualifyRange(ser.values, sheetName));
        xmlStream.closeNode();
        xmlStream.closeNode();
      }

      xmlStream.closeNode(); // c:ser
    });

    if (model.dataLabels || isPieLike) {
      const dl = typeof model.dataLabels === 'object' ? model.dataLabels : {};
      const showVal = dl.showVal !== undefined ? dl.showVal : true;
      const showPercent = dl.showPercent !== undefined ? dl.showPercent : isPieLike;
      const showCatName = dl.showCatName !== undefined ? dl.showCatName : false;
      const showSerName = dl.showSerName !== undefined ? dl.showSerName : false;
      const showLeaderLines = dl.showLeaderLines !== undefined ? dl.showLeaderLines : isPieLike;

      if (model.dataLabels !== false) {
        xmlStream.openNode('c:dLbls');
        xmlStream.leafNode('c:showLegendKey', {val: '0'});
        xmlStream.leafNode('c:showVal', {val: showVal ? '1' : '0'});
        xmlStream.leafNode('c:showCatName', {val: showCatName ? '1' : '0'});
        xmlStream.leafNode('c:showSerName', {val: showSerName ? '1' : '0'});
        xmlStream.leafNode('c:showPercent', {val: showPercent ? '1' : '0'});
        xmlStream.leafNode('c:showBubbleSize', {val: '0'});
        if (dl.separator) {
          xmlStream.leafNode('c:separator', null, String(dl.separator));
        }
        if (isPieLike) {
          xmlStream.leafNode('c:showLeaderLines', {val: showLeaderLines ? '1' : '0'});
        }
        xmlStream.closeNode();
      }
    }

    if (type === 'doughnut') {
      xmlStream.leafNode('c:firstSliceAng', {val: '0'});
      xmlStream.leafNode('c:holeSize', {val: '50'});
    } else if (type === 'pie') {
      xmlStream.leafNode('c:firstSliceAng', {val: '0'});
    } else {
      xmlStream.leafNode('c:axId', {val: String(catAxisId)});
      xmlStream.leafNode('c:axId', {val: String(valAxisId)});
    }

    xmlStream.closeNode(); // chart-type node

    if (!isPieLike) {
      // Category Axis
      xmlStream.openNode('c:catAx');
      xmlStream.leafNode('c:axId', {val: String(catAxisId)});
      xmlStream.openNode('c:scaling');
      xmlStream.leafNode('c:orientation', {val: 'minMax'});
      xmlStream.closeNode();
      xmlStream.leafNode('c:delete', {val: '0'});
      xmlStream.leafNode('c:axPos', {val: type === 'bar' ? 'l' : 'b'});
      xmlStream.leafNode('c:tickLblPos', {val: 'nextTo'});
      xmlStream.leafNode('c:crossAx', {val: String(valAxisId)});
      xmlStream.leafNode('c:crosses', {val: 'autoZero'});
      xmlStream.closeNode();

      // Value Axis
      xmlStream.openNode('c:valAx');
      xmlStream.leafNode('c:axId', {val: String(valAxisId)});
      xmlStream.openNode('c:scaling');
      xmlStream.leafNode('c:orientation', {val: 'minMax'});
      xmlStream.closeNode();
      xmlStream.leafNode('c:delete', {val: '0'});
      xmlStream.leafNode('c:axPos', {val: type === 'bar' ? 'b' : 'l'});
      xmlStream.openNode('c:majorGridlines');
      xmlStream.openNode('c:spPr');
      xmlStream.openNode('a:ln', {w: '9525'});
      xmlStream.openNode('a:solidFill');
      xmlStream.leafNode('a:srgbClr', {val: 'E2E8F0'});
      xmlStream.closeNode();
      xmlStream.closeNode();
      xmlStream.closeNode();
      xmlStream.closeNode();
      if (valAxisNumFmt) {
        xmlStream.leafNode('c:numFmt', {formatCode: valAxisNumFmt, sourceLinked: '0'});
      }
      xmlStream.leafNode('c:majorTickMark', {val: 'none'});
      xmlStream.leafNode('c:minorTickMark', {val: 'none'});
      xmlStream.leafNode('c:tickLblPos', {val: 'nextTo'});
      if (valAxisMajorUnit) {
        xmlStream.leafNode('c:majorUnit', {val: String(valAxisMajorUnit)});
      }
      xmlStream.leafNode('c:crossAx', {val: String(catAxisId)});
      xmlStream.leafNode('c:crosses', {val: 'autoZero'});
      xmlStream.closeNode();
    }

    xmlStream.closeNode(); // c:plotArea

    if (legendPos !== 'none') {
      xmlStream.openNode('c:legend');
      xmlStream.leafNode('c:legendPos', {val: legendPos});
      xmlStream.leafNode('c:overlay', {val: '0'});
      xmlStream.closeNode();
    }

    xmlStream.leafNode('c:plotVisOnly', {val: '1'});
    xmlStream.closeNode(); // c:chart
    xmlStream.closeNode(); // c:chartSpace
  }
}

ChartXform.CHART_SPACE_ATTRIBUTES = {
  'xmlns:c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
  'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
};

module.exports = ChartXform;
