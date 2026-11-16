# 5. JSON Import/Export & Auto-Fit

## `ws.fromJSON(records, options)`

Quickly populate a worksheet from an array of JavaScript objects without manually defining `columns` or looping `addRow()`:

```js
const ws = workbook.addWorksheet('Users');

ws.fromJSON(
  [
    { name: 'Alice', role: 'Admin', score: 98, note: 'Lead Engineer\nJakarta' },
    { name: 'Bob', role: 'Finance', score: 91, note: 'Payroll & Audit' },
  ],
  {
    headers: ['name', 'role', 'score', 'note'], // optional custom order
    headerLabels: { name: 'Full Name', role: 'Role', score: 'Score', note: 'Notes' },
    headerStyle: {
      font: { bold: true, color: { argb: 'FFFFFFFF' } },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F766E' } },
    },
    autoFit: true, // automatically calls ws.autoFitColumns()
  }
);
```

---

## `ws.toJSON(options)`

Convert a worksheet back into an array of plain objects (automatically unwrapping formula results, rich text, and hyperlinks, while guarding against prototype pollution):

```js
const rows = ws.toJSON({
  headerRow: 1,       // set to false to use column letters ('A', 'B', ...)
  startRow: 2,        // defaults to headerRow + 1
  endRow: ws.rowCount,
  includeEmpty: false,
  raw: false,         // true returns raw ExcelJS cell objects (formula/richText)
});
```

---

## `ws.autoFitColumns(options)` & `ws.autoFitRows(options)`

```js
// Auto-fit column widths (supports CJK/Unicode, bold scaling, and font sizes)
ws.autoFitColumns({
  minWidth: 10,
  maxWidth: 60,
  padding: 2,
  onlyEmptyWidth: false,
});

// Auto-fit row heights (supports '\n', wrapText line estimation, and font sizes)
ws.autoFitRows({
  minHeight: 16,
  maxHeight: 409,
  padding: 4,
  onlyEmptyHeight: false,
});
```
