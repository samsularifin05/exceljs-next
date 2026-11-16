const WorksheetWriter = verquire('stream/xlsx/worksheet-writer');
const StreamBuf = verquire('utils/stream-buf');

describe('Workbook Writer', () => {
  it('generates valid xml even when there is no data', () =>
    // issue: https://github.com/guyonroche/exceljs/issues/99
    // PR: https://github.com/guyonroche/exceljs/pull/255
    new Promise((resolve, reject) => {
      const mockWorkbook = {
        _openStream() {
          return this.stream;
        },
        stream: new StreamBuf(),
      };
      mockWorkbook.stream.on('finish', () => {
        try {
          const xml = mockWorkbook.stream.read().toString();
          expect(xml).xml.to.be.valid();
          resolve();
        } catch (error) {
          reject(error);
        }
      });

      const writer = new WorksheetWriter({
        id: 1,
        workbook: mockWorkbook,
      });

      writer.commit();
    }));

  it('finds a cell in uncommitted row via findCell', () => {
    const mockWorkbook = {
      _openStream() {
        return new StreamBuf();
      },
      styles: {addStyleModel: () => 0},
    };
    const writer = new WorksheetWriter({id: 1, workbook: mockWorkbook});
    writer.getCell('B2').value = 'hello';
    expect(writer.findCell('B2').value).to.equal('hello');
    expect(writer.findCell(2, 2).value).to.equal('hello');
  });

  it('supports addTable, addImage, extLst conditional formatting, and WorksheetReader mergeCells in streaming mode', async () => {
    const ExcelJS = verquire('exceljs');
    const outBuf = new StreamBuf();
    const wbWriter = new ExcelJS.stream.xlsx.WorkbookWriter({
      stream: outBuf,
      useStyles: true,
      useSharedStrings: true,
    });

    const imgId = wbWriter.addImage({
      buffer: Buffer.from(
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
        'base64'
      ),
      extension: 'png',
    });

    const wsWriter = wbWriter.addWorksheet('StreamSheet');
    wsWriter.addTable({
      name: 'StreamTable',
      ref: 'A1',
      headerRow: true,
      columns: [{name: 'Col1'}, {name: 'Col2'}],
      rows: [
        ['val1', 10],
        ['val2', 20],
      ],
    });
    expect(wsWriter.getTables().length).to.equal(1);

    wsWriter.addImage(imgId, 'D1:E4');
    expect(wsWriter.getImages().length).to.equal(1);

    wsWriter.addConditionalFormatting({
      ref: 'B2:B3',
      rules: [{type: 'dataBar', priority: 1}],
    });

    wsWriter.mergeCells('A5:B5');
    wsWriter.commit();
    await wbWriter.commit();

    const buffer = outBuf.toBuffer();

    // Verify with in-memory Workbook.xlsx.load
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const ws = wb.getWorksheet('StreamSheet');
    expect(ws.getTables().length).to.equal(1);
    expect(ws.getTable('StreamTable').name).to.equal('StreamTable');
    expect(ws.getImages().length).to.equal(1);
    expect(ws.conditionalFormattings.length).to.equal(1);
    expect(ws.conditionalFormattings[0].rules[0].type).to.equal('dataBar');
    expect(ws.conditionalFormattings[0].rules[0].x14Id).to.be.a('string');

    // Verify WorksheetReader reads mergeCells
    const inStream = new StreamBuf();
    inStream.end(buffer);
    const wbReader = new ExcelJS.stream.xlsx.WorkbookReader(inStream, {
      worksheets: 'emit',
    });
    const merges = [];
    for await (const sheetReader of wbReader) {
      sheetReader.on('mergeCell', ref => merges.push(ref));
      await sheetReader.read();
      expect(sheetReader.mergeCells).to.deep.equal(['A5:B5']);
    }
    expect(merges).to.deep.equal(['A5:B5']);
  });
});
