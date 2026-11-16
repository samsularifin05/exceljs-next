# 11. CSV & Browser Usage

## Reading & Writing CSV

```js
const ExcelJS = require('exceljs-next');

const wb = new ExcelJS.Workbook();
const ws = wb.addWorksheet('Export');
ws.addRows([
  ['Name', 'Score', 'Date'],
  ['Alice', 95, new Date('2026-03-01T00:00:00Z')],
]);

// Write CSV
await wb.csv.writeFile('export.csv', {
  dateFormat: 'YYYY-MM-DD',
  dateUTC: true,
});

// Read CSV
const wbRead = new ExcelJS.Workbook();
await wbRead.csv.readFile('export.csv', {
  dateFormats: ['YYYY-MM-DD', 'MM-DD-YYYY'],
});
```

---

## Browser Usage (`dist/exceljs.min.js`)

`exceljs-next` ships a pre-built browser bundle at `dist/exceljs.min.js` supporting workbook generation, styling, charts, templates, JSON helpers, and even **MS-OFFCRYPTO password encryption/decryption** directly in the browser:

```html
<script src="dist/exceljs.min.js"></script>
<script>
  async function downloadExcel() {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Report');
    ws.fromJSON([{ item: 'Laptop', qty: 2 }], { autoFit: true });

    const buffer = await wb.xlsx.writeBuffer({ password: 'optional-password' });
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'report.xlsx';
    a.click();
    URL.revokeObjectURL(url);
  }
</script>
```
