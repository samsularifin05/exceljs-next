const ExcelJS = verquire('exceljs');

describe('Worksheet JSON & autoFitRows helpers', () => {
  it('converts records via fromJSON and reads back via toJSON', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Users');

    ws.fromJSON(
      [
        {name: 'Alice', role: 'Admin', score: 95},
        {name: 'Bob', role: 'User', score: 80},
      ],
      {
        headerLabels: {name: 'Full Name', role: 'Role', score: 'Score'},
        headerStyle: {font: {bold: true}},
        autoFit: true,
      }
    );

    expect(ws.getCell('A1').value).to.equal('Full Name');
    expect(ws.getCell('A1').font.bold).to.equal(true);
    expect(ws.getColumn(1).width).to.be.greaterThanOrEqual(9);

    const json = ws.toJSON();
    expect(json).to.deep.equal([
      {'Full Name': 'Alice', Role: 'Admin', Score: 95},
      {'Full Name': 'Bob', Role: 'User', Score: 80},
    ]);
  });

  it('ignores unsafe prototype keys in toJSON and fromJSON', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Safe');
    ws.addRow(['__proto__', 'name']);
    ws.addRow(['polluted', 'SafeUser']);

    const rows = ws.toJSON();
    expect(rows).to.deep.equal([{A: 'polluted', name: 'SafeUser'}]);
    expect({}.polluted).to.equal(undefined);
  });

  it('calculates row heights with autoFitRows for multiline and wrapped text', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Heights');
    ws.getColumn(1).width = 10;

    ws.getCell('A1').value = 'Short';
    ws.getCell('A2').value = 'Line 1\nLine 2\nLine 3';
    ws.getCell('A3').value = 'This is a very long wrapped text that spans multiple lines';
    ws.getCell('A3').alignment = {wrapText: true};

    ws.autoFitRows({minHeight: 16, padding: 4});

    expect(ws.getRow(1).height).to.be.greaterThanOrEqual(16);
    expect(ws.getRow(2).height).to.be.greaterThan(ws.getRow(1).height);
    expect(ws.getRow(3).height).to.be.greaterThan(ws.getRow(2).height);
  });
});
