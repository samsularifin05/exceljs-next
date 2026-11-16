const JSZip = require('jszip');
const ExcelJS = verquire('exceljs');

describe('Worksheet Native Charts (addChart)', () => {
  it('generates valid chart XML, drawing anchor, relationships, and Content_Types', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sales');
    ws.addRows([
      ['Month', 'Revenue', 'Target'],
      ['Jan', 120, 100],
      ['Feb', 150, 130],
      ['Mar', 180, 160],
    ]);

    ws.addChart({
      type: 'column',
      title: 'Monthly Revenue vs Target',
      range: 'E2:L16',
      series: [
        {name: 'Revenue', categories: '$A$2:$A$4', values: '$B$2:$B$4', color: '#10B981'},
        {name: 'Target', categories: '$A$2:$A$4', values: '$C$2:$C$4', color: '#3B82F6'},
      ],
    });

    ws.addChart({
      type: 'pie',
      title: 'Share',
      range: 'E18:L30',
      series: [{name: 'Revenue', categories: '$A$2:$A$4', values: '$B$2:$B$4'}],
    });

    expect(ws.getCharts().length).to.equal(2);

    const buf = await wb.xlsx.writeBuffer();
    const zip = await JSZip.loadAsync(buf);

    const chart1Xml = await zip.file('xl/charts/chart1.xml').async('string');
    const chart2Xml = await zip.file('xl/charts/chart2.xml').async('string');
    const drawingXml = await zip.file('xl/drawings/drawing1.xml').async('string');
    const drawingRels = await zip.file('xl/drawings/_rels/drawing1.xml.rels').async('string');
    const contentTypes = await zip.file('[Content_Types].xml').async('string');

    expect(chart1Xml).to.include('<c:barChart>');
    expect(chart1Xml).to.include('<c:barDir val="col"/>');
    expect(chart1Xml).to.include('Monthly Revenue vs Target');
    expect(chart1Xml).to.include('&apos;Sales&apos;!$A$2:$A$4');
    expect(chart2Xml).to.include('<c:pieChart>');
    expect(chart2Xml).to.include('<c:dLbls>');
    expect(chart2Xml).to.include('<c:showPercent val="1"/>');
    expect(drawingXml).to.include('<c:chart');
    expect(drawingRels).to.include('../charts/chart1.xml');
    expect(drawingRels).to.include('../charts/chart2.xml');
    expect(contentTypes).to.include('/xl/charts/chart1.xml');
    expect(contentTypes).to.include('/xl/charts/chart2.xml');
  });
});
