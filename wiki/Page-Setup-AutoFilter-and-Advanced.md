# 12. Page Setup, Auto Filters, Outlines, & Defined Names

## Page Setup, Margins, & Print Area

```js
const ws = workbook.addWorksheet('PrintReport', {
  pageSetup: {
    paperSize: 9, // A4
    orientation: 'landscape',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    margins: {
      left: 0.5, right: 0.5,
      top: 0.75, bottom: 0.75,
      header: 0.3, footer: 0.3,
    },
    printArea: 'A1:G50',
    printTitlesRow: '1:2', // Repeat rows 1-2 on every printed page
  },
});
```

---

## Headers & Footers

Supports special codes like `&P` (current page), `&N` (total pages), `&D` (date), `&T` (time), `&B` (bold), `&F` (filename):

```js
ws.headerFooter.oddHeader = '&C&BMonthly Sales Report';
ws.headerFooter.oddFooter = '&LConfidential&RPage &P of &N';
```

---

## Auto Filters

```js
// Enable filter dropdowns on A1:D1
ws.autoFilter = 'A1:D1';
// Or using row/col coordinates
ws.autoFilter = {
  from: { row: 1, column: 1 },
  to: { row: 1, column: 4 },
};
```

---

## Outline Levels (Row & Column Grouping)

```js
// Group rows 2..4 under outline level 1
ws.getRow(2).outlineLevel = 1;
ws.getRow(3).outlineLevel = 1;
ws.getRow(4).outlineLevel = 1;

// Collapse/expand threshold
ws.properties.outlineLevelRow = 1;
```

---

## Defined Names (Named Ranges)

```js
// Name a cell or range for use in formulas
ws.getCell('B2').name = 'TaxRate';
ws.getCell('C2').value = { formula: 'A2*TaxRate' };

// Or add via workbook.definedNames
workbook.definedNames.add('Sheet1!$A$2:$A$10', 'PriceList');
```

---

## Value Types (`ExcelJS.ValueType`)

Check what type of value a cell holds:

```js
const { ValueType } = require('exceljs-next');

const cell = ws.getCell('A1');
if (cell.type === ValueType.Formula) {
  console.log('Formula:', cell.formula, 'Result:', cell.result);
} else if (cell.type === ValueType.Date) {
  console.log('Date:', cell.value);
}
```
