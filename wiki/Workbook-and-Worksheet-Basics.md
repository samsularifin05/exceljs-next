# 7. Workbook & Worksheet Basics

## Creating & Configuring a Workbook

```js
const ExcelJS = require('exceljs-next');

const workbook = new ExcelJS.Workbook();
workbook.creator = 'Engineering Team';
workbook.lastModifiedBy = 'API Service';
workbook.created = new Date();
workbook.modified = new Date();
```

---

## Managing Worksheets

```js
// Add worksheet with frozen header row
const ws = workbook.addWorksheet('Sales', {
  properties: { tabColor: { argb: 'FF10B981' }, defaultColWidth: 15 },
  views: [{ state: 'frozen', ySplit: 1 }],
});

// Access worksheets
const byName = workbook.getWorksheet('Sales');
const byId = workbook.getWorksheet(1);

// Iterate all sheets
workbook.eachSheet((worksheet, sheetId) => {
  console.log(sheetId, worksheet.name);
});
```

---

## Columns, Rows, & Cells

```js
ws.columns = [
  { header: 'ID', key: 'id', width: 10 },
  { header: 'Name', key: 'name', width: 28 },
  { header: 'Joined', key: 'joined', width: 16 },
];

ws.addRow({ id: 1, name: 'Alice', joined: new Date(2026, 0, 15) });
ws.addRows([
  [2, 'Bob', new Date(2026, 1, 20)],
  [3, 'Charlie', new Date(2026, 2, 10)],
]);

// Update individual cell
ws.getCell('B2').value = 'Alice Wijaya';
```

---

## Rich Text, Hyperlinks, & Merged Cells

```js
// Rich Text
ws.getCell('A5').value = {
  richText: [
    { font: { bold: true, color: { argb: 'FF10B981' } }, text: 'Status: ' },
    { font: { italic: true }, text: 'Approved' },
  ],
};

// Hyperlink
ws.getCell('A6').value = {
  text: 'Open Dashboard',
  hyperlink: 'https://example.com/dashboard',
  tooltip: 'Click to open',
};

// Merge & Unmerge Cells
ws.mergeCells('A8:C8');
ws.getCell('A8').value = 'Merged Header';
// ws.unMergeCells('A8:C8');
```
