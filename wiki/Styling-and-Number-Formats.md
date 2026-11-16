# 8. Styling & Number Formats

## Fonts, Fills, Alignment, & Borders

```js
const cell = ws.getCell('A1');

// Font
cell.font = {
  name: 'Calibri',
  size: 12,
  bold: true,
  italic: false,
  underline: false,
  color: { argb: 'FFFFFFFF' },
};

// Solid Pattern Fill
cell.fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FF0F766E' },
};

// Alignment & Text Wrapping
cell.alignment = {
  vertical: 'middle',
  horizontal: 'center',
  wrapText: true,
  shrinkToFit: false,
  textRotation: 0,
};

// Borders
cell.border = {
  top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
  left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
  bottom: { style: 'medium', color: { argb: 'FF0F766E' } },
  right: { style: 'thin', color: { argb: 'FFCBD5E1' } },
};
```

---

## Gradient Fills

```js
ws.getCell('B1').fill = {
  type: 'gradient',
  gradient: 'angle',
  degree: 90,
  stops: [
    { position: 0, color: { argb: 'FF10B981' } },
    { position: 1, color: { argb: 'FF047857' } },
  ],
};
```

---

## Number Formats (`numFmt`)

```js
ws.getCell('C2').numFmt = '"Rp"#,##0';      // Currency (IDR)
ws.getCell('D2').numFmt = '$#,##0.00';      // Currency (USD)
ws.getCell('E2').numFmt = '0.0%';           // Percentage
ws.getCell('F2').numFmt = 'yyyy-mm-dd';     // Date
ws.getCell('G2').numFmt = '#,##0.00';       // Decimal
```

---

## Row & Column Style Inheritance

Apply styles to an entire row or column at once:

```js
ws.getRow(1).font = { bold: true };
ws.getRow(1).height = 24;
ws.getColumn(3).numFmt = '"Rp"#,##0';
```
