# 2. File Encryption & Protection

Excel / WPS Office have **three distinct protection mechanisms**. Choose the right one for your use case:

| Feature | Method | What It Protects | Prompts Password on File Open? |
| --- | --- | --- | --- |
| **File Encryption** | `wb.xlsx.writeFile(path, { password })` | Encrypts the entire `.xlsx` file (`MS-OFFCRYPTO` AES-256) | **Yes** |
| **Worksheet Protection** | `await ws.protect(password, options)` | Locks cells from editing, formatting, inserting/deleting rows | No (only when editing locked cells / unprotecting sheet) |
| **Workbook Structure Protection** | `await wb.protect(password, options)` | Locks sheet tabs (prevents adding, deleting, renaming, unhiding sheets) | No (only when clicking *Unprotect Workbook*) |

---

## 1. File Encryption (Password to Open `.xlsx`)

Uses **MS-OFFCRYPTO Agile Encryption (AES-256-CBC + SHA-512 + HMAC-SHA512)** on write, and supports decrypting both **Agile (`4.4`)** and **Standard (`3.2`/`4.2`)** encrypted `.xlsx` files on read.

```js
const ExcelJS = require('exceljs-next');

// Write password-encrypted .xlsx
const wb = new ExcelJS.Workbook();
wb.addWorksheet('Payroll').addRow(['Name', 'Salary']);
await wb.xlsx.writeFile('payroll.xlsx', { password: 'Secret@2026' });

// Also works with writeBuffer() and write(stream)
const encryptedBuffer = await wb.xlsx.writeBuffer({ password: 'Secret@2026' });

// Read / Decrypt .xlsx
const wb2 = new ExcelJS.Workbook();
await wb2.xlsx.readFile('payroll.xlsx', { password: 'Secret@2026' });
// Or from buffer:
await wb2.xlsx.load(encryptedBuffer, { password: 'Secret@2026' });
```

---

## 2. Worksheet Protection (Lock Editing Cells)

Prevents users from editing cells unless explicitly unlocked:

```js
const ws = wb.addWorksheet('Form');
ws.getCell('A1').value = 'Read-only label';
ws.getCell('B1').value = 'Editable input';

// Allow user to edit ONLY cell B1
ws.getCell('B1').protection = { locked: false };

await ws.protect('sheet-pass', {
  selectLockedCells: true,
  selectUnlockedCells: true,
  formatCells: false,
  insertRows: false,
  deleteRows: false,
});
```

---

## 3. Workbook Structure Protection (Lock Sheet Tabs)

Prevents users from adding, deleting, renaming, moving, or unhiding sheets (useful when you have hidden lookup/config sheets):

```js
await wb.protect('structure-pass', {
  lockStructure: true,
  lockWindows: false,
});

// Remove protection from workbook model
wb.unprotect();
```
