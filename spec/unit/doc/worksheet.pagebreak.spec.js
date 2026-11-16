const Excel = verquire('exceljs');

describe('Worksheet', () => {
  describe('Page Breaks', () => {
    it('adds multiple row breaks', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('blort');

      // initial values
      ws.getCell('A1').value = 'A1';
      ws.getCell('B1').value = 'B1';
      ws.getCell('A2').value = 'A2';
      ws.getCell('B2').value = 'B2';
      ws.getCell('A3').value = 'A3';
      ws.getCell('B3').value = 'B3';

      let row = ws.getRow(1);
      row.addPageBreak();
      row = ws.getRow(2);
      row.addPageBreak();

      expect(ws.rowBreaks.length).to.equal(2);
    });

    it('adds multiple column breaks and round-trips through xlsx', async () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('blort');

      ws.getCell('A1').value = 'A1';
      ws.getCell('B1').value = 'B1';
      ws.getCell('C1').value = 'C1';

      ws.getColumn(1).addPageBreak();
      ws.getColumn('B').addPageBreak(2, 10);
      ws.getRow(1).addPageBreak();

      expect(ws.colBreaks).to.deep.equal([
        {id: 1, max: 1048575, man: 1},
        {id: 2, min: 1, max: 9, man: 1},
      ]);

      const buffer = await wb.xlsx.writeBuffer();
      const wb2 = new Excel.Workbook();
      await wb2.xlsx.load(buffer);
      const ws2 = wb2.getWorksheet('blort');

      expect(ws2.colBreaks).to.deep.equal([
        {id: 1, max: 1048575, man: 1},
        {id: 2, min: 1, max: 9, man: 1},
      ]);
      expect(ws2.rowBreaks).to.deep.equal([{id: 1, max: 16838, man: 1}]);
    });
  });
});
