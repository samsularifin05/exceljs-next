const ExcelJS = verquire('exceljs');
const offcrypto = verquire('utils/offcrypto');

describe('offcrypto (File Encryption / Decryption)', () => {
  it('encrypts and decrypts an xlsx workbook buffer with password', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Payroll');
    ws.addRow(['Employee', 'Salary']);
    ws.addRow(['Budi', 15000000]);

    const encrypted = await wb.xlsx.writeBuffer({password: 'Secret@2026'});
    expect(offcrypto.isCfbBuffer(encrypted)).to.equal(true);

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(encrypted, {password: 'Secret@2026'});
    const ws2 = wb2.getWorksheet('Payroll');
    expect(ws2.getCell('A2').value).to.equal('Budi');
    expect(ws2.getCell('B2').value).to.equal(15000000);
  });

  it('rejects wrong password or missing password on encrypted workbook', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('Sheet1').getCell('A1').value = 'TopSecret';
    const encrypted = await wb.xlsx.writeBuffer({password: 'correct-pass'});

    const wbWrong = new ExcelJS.Workbook();
    let wrongErr = null;
    try {
      await wbWrong.xlsx.load(encrypted, {password: 'wrong-pass'});
    } catch (e) {
      wrongErr = e;
    }
    expect(wrongErr).to.be.an('error');
    expect(wrongErr.message).to.include('Incorrect password');

    const wbMissing = new ExcelJS.Workbook();
    let missingErr = null;
    try {
      await wbMissing.xlsx.load(encrypted);
    } catch (e) {
      missingErr = e;
    }
    expect(missingErr).to.be.an('error');
    expect(missingErr.message).to.include('password');
  });

  it('handles multi-segment payloads (> 32KB) cleanly', () => {
    const payload = Buffer.alloc(40000, 0x41);
    const cfb = offcrypto.encryptWorkbookBuffer(payload, 'multi-seg');
    expect(offcrypto.isCfbBuffer(cfb)).to.equal(true);
    const decrypted = offcrypto.decryptWorkbookBuffer(cfb, 'multi-seg');
    expect(decrypted.equals(payload)).to.equal(true);
  });
});
