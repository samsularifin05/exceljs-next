# 3. Native Excel Charts & Data Labels

Create native OOXML Excel charts directly inside any worksheet using `ws.addChart(options)`.

## Supported Chart Types
- `'column'` (Vertical bar chart — default)
- `'bar'` (Horizontal bar chart)
- `'line'` (Line chart)
- `'area'` (Area chart)
- `'pie'` (Pie chart)
- `'doughnut'` (Doughnut chart)

---

## Example: Column Chart with Formatted Axis

```js
const wb = new ExcelJS.Workbook();
const ws = wb.addWorksheet('Sales');

ws.addRows([
  ['Month', 'Revenue', 'Target'],
  ['Jan', 12000000, 10000000],
  ['Feb', 15000000, 13000000],
  ['Mar', 21000000, 16000000],
]);

ws.addChart({
  type: 'column',
  title: 'Monthly Revenue vs Target',
  range: 'E2:N18',
  legend: { position: 'b' }, // 'r' | 'b' | 't' | 'l' | 'none'
  valAxis: {
    numFmt: '#,##0,," Jt"', // Prevent overlapping long numbers
  },
  series: [
    { name: 'Revenue', categories: '$A$2:$A$4', values: '$B$2:$B$4', color: '#10B981' },
    { name: 'Target', categories: '$A$2:$A$4', values: '$C$2:$C$4', color: '#3B82F6' },
  ],
});
```

---

## Example: Doughnut / Pie Chart with Percentage Data Labels

For `'pie'` and `'doughnut'` charts, `dataLabels` defaults to showing both values and percentages (`showVal: true`, `showPercent: true`). You can customize it via `dataLabels`:

```js
ws.addChart({
  type: 'doughnut',
  title: 'Revenue Share',
  range: 'E2:M18',
  legend: { position: 'b' },
  dataLabels: {
    showVal: true,
    showPercent: true,
    showCatName: false,
    showLeaderLines: true,
    separator: '\n',
  },
  series: [
    {
      name: 'Revenue',
      categories: '$A$2:$A$4',
      values: '$B$2:$B$4',
    },
  ],
});
```

*(Pass `dataLabels: false` if you want to hide labels on a pie/doughnut chart.)*
