# 1. Installation & Quick Start

## Requirements
- **Node.js:** `>= 16.0.0` (tested on Node 16, 18, 20, and 22+)
- **Browsers:** Modern browsers via `dist/exceljs.min.js`

---

## Installation

### Standard Install
```bash
npm install exceljs-next
# or with pnpm
pnpm add exceljs-next
```

### Drop-In Replacement for `exceljs`
If your existing codebase already uses `require('exceljs')` or `import ExcelJS from 'exceljs'`, install `exceljs-next` as an alias so you don't need to edit any imports:

```bash
npm install exceljs@npm:exceljs-next
```

---

## Importing in Your Project

### CommonJS (Node.js)
```js
const ExcelJS = require('exceljs-next');

const workbook = new ExcelJS.Workbook();
```

### Native ESM (`.mjs` or `"type": "module"`)
```js
import ExcelJS, { Workbook, ValueType } from 'exceljs-next';

const workbook = new Workbook();
```

### TypeScript
`exceljs-next` ships both `index.d.ts` (CommonJS) and `index.d.mts` (ESM), compatible with `strict: true` and `moduleResolution`: `"node"`, `"node16"`, `"nodenext"`, or `"bundler"`:

```ts
import ExcelJS, { Workbook, Worksheet } from 'exceljs-next';

const wb: Workbook = new ExcelJS.Workbook();
const ws: Worksheet = wb.addWorksheet('Report');
ws.getCell('A1').value = 'Hello TypeScript';
await wb.xlsx.writeFile('hello.xlsx');
```
