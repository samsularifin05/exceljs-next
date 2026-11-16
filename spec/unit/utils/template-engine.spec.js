const ExcelJS = verquire('exceljs');

describe('Template Engine (renderTemplate)', () => {
  it('replaces scalar placeholders, preserves native types, and expands array loop rows', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Invoice');

    ws.getCell('A1').value = 'Invoice #{{invoiceNo}} - {{customer.name}}';
    ws.getCell('B2').value = '{{issuedDate}}';
    ws.getCell('A4').value = 'Item';
    ws.getCell('B4').value = 'Qty';
    ws.getCell('C4').value = 'Price';
    ws.getCell('A5').value = '{{items.$number}}. {{items.name}}';
    ws.getCell('B5').value = '{{items.qty}}';
    ws.getCell('C5').value = '{{items.price}}';
    ws.getCell('A5').font = {bold: true};

    const issuedDate = new Date('2026-03-10T00:00:00Z');
    wb.renderTemplate({
      invoiceNo: 'INV-001',
      customer: {name: 'PT Maju Jaya'},
      issuedDate,
      items: [
        {name: 'Laptop', qty: 2, price: 15000000},
        {name: 'Mouse', qty: 5, price: 250000},
      ],
    });

    expect(ws.getCell('A1').value).to.equal('Invoice #INV-001 - PT Maju Jaya');
    expect(ws.getCell('B2').value).to.deep.equal(issuedDate);
    expect(ws.getCell('A5').value).to.equal('1. Laptop');
    expect(ws.getCell('B5').value).to.equal(2);
    expect(ws.getCell('C5').value).to.equal(15000000);
    expect(ws.getCell('A6').value).to.equal('2. Mouse');
    expect(ws.getCell('B6').value).to.equal(5);
    expect(ws.getCell('C6').value).to.equal(250000);
    expect(ws.getCell('A6').font.bold).to.equal(true);
  });

  it('removes empty loop rows and ignores unsafe prototype keys', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Empty');
    ws.getCell('A1').value = '{{__proto__.polluted}}';
    ws.getCell('A2').value = '{{items.name}}';
    ws.getCell('A3').value = 'Footer';

    ws.renderTemplate({items: []});
    expect(ws.getCell('A1').value).to.equal('');
    expect(ws.getCell('A2').value).to.equal('Footer');
    expect({}.polluted).to.equal(undefined);
  });
});
