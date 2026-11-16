// *Note*: `fs.promises` not supported before Node.js 11.14.0;
// ExcelJS version range '>=8.3.0' (as of 2023-10-08).
const fs = require('fs');
const {promisify} = require('util');

const fsReadFileAsync = promisify(fs.readFile);

const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

const PIVOT_TABLE_FILEPATHS = [
  'xl/pivotCache/pivotCacheRecords1.xml',
  'xl/pivotCache/pivotCacheDefinition1.xml',
  'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels',
  'xl/pivotTables/pivotTable1.xml',
  'xl/pivotTables/_rels/pivotTable1.xml.rels',
];

const TEST_XLSX_FILEPATH = './spec/out/wb.test.xlsx';

const TEST_DATA = [
  ['A', 'B', 'C', 'D', 'E'],
  ['a1', 'b1', 'c1', 4, 5],
  ['a1', 'b2', 'c1', 4, 5],
  ['a2', 'b1', 'c2', 14, 24],
  ['a2', 'b2', 'c2', 24, 35],
  ['a3', 'b1', 'c3', 34, 45],
  ['a3', 'b2', 'c3', 44, 45],
];

// =============================================================================
// Tests

describe('Workbook', () => {
  describe('Pivot Tables', () => {
    it('if pivot table added, then certain xml and rels files are added', async () => {
      const workbook = new ExcelJS.Workbook();

      const worksheet1 = workbook.addWorksheet('Sheet1');
      worksheet1.addRows(TEST_DATA);

      const worksheet2 = workbook.addWorksheet('Sheet2');
      worksheet2.addPivotTable({
        sourceSheet: worksheet1,
        rows: ['A', 'B'],
        columns: ['C'],
        values: ['E'],
        metric: 'sum',
      });

      return workbook.xlsx.writeFile(TEST_XLSX_FILEPATH).then(async () => {
        const buffer = await fsReadFileAsync(TEST_XLSX_FILEPATH);
        const zip = await JSZip.loadAsync(buffer);
        for (const filepath of PIVOT_TABLE_FILEPATHS) {
          expect(zip.files[filepath]).to.not.be.undefined();
        }
      });
    });

    it('if pivot table NOT added, then certain xml and rels files are not added', () => {
      const workbook = new ExcelJS.Workbook();

      const worksheet1 = workbook.addWorksheet('Sheet1');
      worksheet1.addRows(TEST_DATA);

      workbook.addWorksheet('Sheet2');

      return workbook.xlsx.writeFile(TEST_XLSX_FILEPATH).then(async () => {
        const buffer = await fsReadFileAsync(TEST_XLSX_FILEPATH);
        const zip = await JSZip.loadAsync(buffer);
        for (const filepath of PIVOT_TABLE_FILEPATHS) {
          expect(zip.files[filepath]).to.be.undefined();
        }
      });
    });

    it('supports multiple pivot tables, multiple values, custom metrics, and sourceTable', async () => {
      const workbook = new ExcelJS.Workbook();

      const worksheet1 = workbook.addWorksheet('Sheet1');
      const table = worksheet1.addTable({
        name: 'SalesTable',
        ref: 'A1',
        headerRow: true,
        columns: [
          {name: 'A'},
          {name: 'B'},
          {name: 'C'},
          {name: 'D'},
          {name: 'E'},
        ],
        rows: TEST_DATA.slice(1),
      });

      const worksheet2 = workbook.addWorksheet('Sheet2');
      const pt1 = worksheet2.addPivotTable({
        sourceSheet: worksheet1,
        rows: ['A'],
        columns: ['C'],
        values: ['D', 'E'],
        metric: 'count',
      });
      expect(pt1.metric).to.equal('count');
      expect(pt1.tableNumber).to.equal(1);

      const worksheet3 = workbook.addWorksheet('Sheet3');
      const pt2 = worksheet3.addPivotTable({
        sourceTable: table,
        rows: ['B'],
        columns: ['C'],
        values: ['E'],
        metric: 'average',
      });
      expect(pt2.metric).to.equal('average');
      expect(pt2.tableNumber).to.equal(2);

      const buffer = await workbook.xlsx.writeBuffer();
      const zip = await JSZip.loadAsync(buffer);
      expect(zip.files['xl/pivotTables/pivotTable1.xml']).to.not.be.undefined();
      expect(zip.files['xl/pivotTables/pivotTable2.xml']).to.not.be.undefined();
      expect(zip.files['xl/pivotCache/pivotCacheDefinition2.xml']).to.not.be.undefined();
      expect(zip.files['xl/pivotCache/pivotCacheRecords2.xml']).to.not.be.undefined();

      const pt1Xml = await zip.files['xl/pivotTables/pivotTable1.xml'].async('string');
      expect(pt1Xml).to.include('subtotal="count"');
      expect(pt1Xml).to.include('Count of D');
      expect(pt1Xml).to.include('Count of E');

      const pt2Xml = await zip.files['xl/pivotTables/pivotTable2.xml'].async('string');
      expect(pt2Xml).to.include('subtotal="average"');
      expect(pt2Xml).to.include('Average of E');
    });

    it('escapes XML special characters and handles null/undefined values in pivot tables', async () => {
      const workbook = new ExcelJS.Workbook();
      const ws1 = workbook.addWorksheet('Source');
      ws1.addRows([
        ['Company & Co', 'Region', 'Amount'],
        ['A & B <Corp>', 'North', 100],
        ['Mid Null', null, 150],
        ['C "Quote" \'Single\'', 'South', 200],
      ]);

      const ws2 = workbook.addWorksheet('Pivot');
      ws2.addPivotTable({
        sourceSheet: ws1,
        rows: ['Company & Co'],
        columns: ['Region'],
        values: ['Amount'],
        metric: 'sum',
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const zip = await JSZip.loadAsync(buffer);
      const cacheDefXml = await zip.files['xl/pivotCache/pivotCacheDefinition1.xml'].async('string');
      const cacheRecXml = await zip.files['xl/pivotCache/pivotCacheRecords1.xml'].async('string');

      expect(cacheDefXml).to.include('name="Company &amp; Co"');
      expect(cacheDefXml).to.include('v="A &amp; B &lt;Corp&gt;"');
      expect(cacheRecXml).to.include('<m />');
      expect(cacheRecXml).to.not.include('undefined');
    });
  });
});
