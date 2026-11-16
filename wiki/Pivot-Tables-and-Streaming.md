# 6. Pivot Tables & Streaming

## Enhanced Multi-Pivot Tables

`exceljs-next` supports multiple Pivot Tables per workbook, 11 aggregation metrics (`sum`, `count`, `average`, `max`, `min`, `product`, `countNums`, `stdDev`, `stdDevP`, `var`, `varP`), multiple `values` fields, and `sourceTable`:

```js
const pivotSheet = workbook.addWorksheet('Pivot Summary');
pivotSheet.addPivotTable({
  sourceTable: salesTable, // or sourceSheet: dataSheet
  rows: ['Region', 'Product'],
  columns: [],
  values: ['Revenue', 'Units'],
  metric: 'average',
});
```

---

## Streaming Writer (`WorkbookWriter`)

Write hundreds of thousands of rows with minimal memory usage — now with full support for `addTable()`, `addImage()`, extended conditional formatting (`dataBar` / `iconSet` `<extLst>`), and `colBreaks`:

```js
const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
  filename: 'large-export.xlsx',
  useStyles: true,
  useSharedStrings: true,
});

const ws = workbook.addWorksheet('Data');
ws.addTable({
  name: 'BigTable',
  ref: 'A1',
  headerRow: true,
  columns: [{ name: 'ID' }, { name: 'Name' }],
  rows: [[1, 'Alice'], [2, 'Bob']],
});

ws.getColumn(1).addPageBreak(0, 1048575);
ws.commit();
await workbook.commit();
```

---

## Streaming Reader (`WorkbookReader`)

Works reliably on **Node.js 16, 18, 20, and 22+** (including multi-byte UTF-8 chunks and `mergeCells` parsing):

```js
const reader = new ExcelJS.stream.xlsx.WorkbookReader('large-export.xlsx', {
  worksheets: 'emit',
  sharedStrings: 'cache',
});

for await (const wsReader of reader) {
  wsReader.on('mergeCell', range => {
    console.log('Merged range:', range);
  });
  for await (const row of wsReader) {
    console.log(row.values);
  }
}
```
