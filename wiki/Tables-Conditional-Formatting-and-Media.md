# 10. Tables, Conditional Formatting, & Media

## Excel Tables (`ws.addTable()`)

```js
const salesTable = ws.addTable({
  name: 'SalesTable',
  ref: 'A1',
  headerRow: true,
  totalsRow: true,
  style: {
    theme: 'TableStyleMedium9',
    showRowStripes: true,
  },
  columns: [
    { name: 'Product', filterButton: true, totalsRowLabel: 'Total:' },
    { name: 'Units', filterButton: true, totalsRowFunction: 'sum' },
    { name: 'Revenue', filterButton: true, totalsRowFunction: 'sum' },
  ],
  rows: [
    ['Laptop', 5, 75000000],
    ['Monitor', 10, 40000000],
  ],
});
```

---

## Conditional Formatting

```js
ws.addConditionalFormatting({
  ref: 'C2:C10',
  rules: [
    {
      type: 'cellIs',
      operator: 'greaterThan',
      formulae: [10000000],
      style: {
        fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFD1FAE5' } },
        font: { color: { argb: 'FF065F46' }, bold: true },
      },
    },
  ],
});
```

---

## Images, Comments, & Page Breaks

```js
// Add image from file or Buffer
const imageId = workbook.addImage({
  filename: 'logo.png',
  extension: 'png',
});
ws.addImage(imageId, 'E2:G6');

// Add background image
ws.addBackgroundImage(imageId);

// Add cell comment / note
ws.getCell('A2').note = {
  texts: [{ font: { bold: true }, text: 'Verified by Finance' }],
};

// Row & Column Page Breaks
ws.getRow(25).addPageBreak();
ws.getColumn(3).addPageBreak(0, 1048575);
```
