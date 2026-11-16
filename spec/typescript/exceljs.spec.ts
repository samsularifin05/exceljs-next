import 'regenerator-runtime/runtime';
import { PassThrough } from 'stream';
import { expect } from 'chai';
import ExcelJS from '../../index';

describe('typescript', () => {
  it('can create and buffer xlsx', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');
    ws.getCell('A1').value = 7;
    const buffer = await wb.xlsx.writeBuffer({
      useStyles: true,
      useSharedStrings: true,
    });

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(buffer);
    const ws2 = wb2.getWorksheet('blort');
    expect(ws2.getCell('A1').value).to.equal(7);
  });

  it('can create and stream xlsx', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');
    ws.getCell('A1').value = 7;

    const stream = new PassThrough();
    await wb.xlsx.write(stream);
    stream.end();

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.read(stream);
    const ws2 = wb2.getWorksheet('blort');
    expect(ws2.getCell('A1').value).to.equal(7);
  });

  it('types getTables, row dimensions, column values, and addPivotTable accurately', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.addRows([
      ['A', 'B', 'C'],
      ['a1', 'b1', 10],
    ]);

    const rowDims: { min: number; max: number } | null = ws.getRow(1).dimensions;
    expect(rowDims).to.deep.equal({ min: 1, max: 3 });

    const colVals: ReadonlyArray<ExcelJS.CellValue> = ws.getColumn(1).values;
    expect(colVals.length).to.be.greaterThan(0);

    const tables: ExcelJS.Table[] = ws.getTables();
    expect(tables).to.be.an('array');

    const ws2 = wb.addWorksheet('Pivot');
    const pt = ws2.addPivotTable({
      sourceSheet: ws,
      rows: ['A'],
      columns: ['B'],
      values: ['C'],
      metric: 'count',
    });
    expect(pt.metric).to.equal('count');
    expect(pt.tableNumber).to.equal(1);

    ws.autoFitColumns({ minWidth: 10, maxWidth: 50 });
    expect(ws.getColumn(1).width).to.be.greaterThanOrEqual(10);

    ws.getColumn(1).addPageBreak();
    expect(ws.colBreaks.length).to.equal(1);
  });

  it('supports workbook protection in typescript', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('Sheet1');
    await wb.protect('secret', { lockStructure: true });
    expect(wb.workbookProtection?.lockStructure).to.equal(true);
    wb.unprotect();
    expect(wb.workbookProtection).to.equal(null);
  });

  it('supports autoFitRows, toJSON, fromJSON, and password-encrypted xlsx in typescript', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Data');
    ws.fromJSON([{ id: 1, name: 'Item A\nLine 2' }], { autoFit: true });
    ws.autoFitRows({ minHeight: 18 });
    const rows = ws.toJSON<{ id: number; name: string }>();
    expect(rows[0].name).to.equal('Item A\nLine 2');

    const encrypted = await wb.xlsx.writeBuffer({ password: 'ts-secret' });
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(encrypted, { password: 'ts-secret' });
    expect(wb2.getWorksheet('Data').getCell('A2').value).to.equal(1);
  });

  it('supports renderTemplate and addChart in typescript', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Template');
    ws.getCell('A1').value = 'Hello {{user}}';
    ws.getCell('A2').value = '{{items.name}}';
    ws.getCell('B2').value = '{{items.val}}';
    wb.renderTemplate({ user: 'Budi', items: [{ name: 'A', val: 10 }, { name: 'B', val: 20 }] });
    ws.addChart({
      type: 'line',
      title: 'Trend',
      range: 'D2:J14',
      series: [{ name: 'Val', categories: '$A$2:$A$3', values: '$B$2:$B$3' }],
    });
    expect(ws.getCharts().length).to.equal(1);
    const buf = await wb.xlsx.writeBuffer();
    expect(buf.byteLength).to.be.greaterThan(0);
  });
});
