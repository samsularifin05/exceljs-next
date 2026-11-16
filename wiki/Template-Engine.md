# 4. Template Engine (`renderTemplate`)

Load an existing `.xlsx` template (designed in Excel or WPS with your borders, logos, fonts, and formulas) or build one in code, then populate `{{placeholder}}` variables and expand array rows automatically.

---

## Supported Syntax
- **Scalar & Nested Paths:** `{{invoiceNo}}`, `{{customer.name}}`, `{{company.address.city}}`
- **Native Type Preservation:** If a cell contains *only* a single placeholder (e.g. `{{items.qty}}` or `{{issuedDate}}`), `exceljs-next` preserves its native `number`, `Date`, or `boolean` type instead of converting it to a string.
- **Array Row Loops:** Any row containing `{{arrayName.field}}` (or `{{arrayName[].field}}`) is automatically duplicated for each element in `data.arrayName`, preserving cell styles and number formats.
  - Built-in loop helpers: `{{items.$index}}` (0-based) and `{{items.$number}}` (1-based).
  - If the array is empty (`[]`), the template row is removed by default (`removeEmptyLoopRows: true`).

---

## Complete Example

```js
const ExcelJS = require('exceljs-next');

const wb = new ExcelJS.Workbook();
await wb.xlsx.readFile('invoice-template.xlsx');

wb.renderTemplate({
  invoiceNo: 'INV-2026-001',
  issuedDate: new Date('2026-03-15'),
  customer: {
    name: 'PT Nusantara Digital',
    city: 'Jakarta',
  },
  items: [
    { name: 'Cloud Server Pro', qty: 2, price: 4500000 },
    { name: 'Managed Database', qty: 1, price: 3200000 },
    { name: 'SSL Wildcard', qty: 3, price: 850000 },
  ],
});

await wb.xlsx.writeFile('invoice-filled.xlsx');
```

You can also render a single worksheet with `worksheet.renderTemplate(data, options)`.
