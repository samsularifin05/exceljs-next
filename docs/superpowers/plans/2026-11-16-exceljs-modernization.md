# ExcelJS Modernization & Security Hardening Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fix all security vulnerabilities (prototype pollution & dependency advisories), fix streaming reader & runtime bugs on Node 18+, remove obsolete Node 8/10 workarounds and duplicated code, and bring TypeScript definitions (`index.d.ts`) and tests into full sync without changing existing logic.

**Architecture:** Maintain the existing CommonJS + Browserify architecture (`lib/` -> `build/` -> `dist/`) so all Node and browser consumers continue to work identically. Harden object merges/maps against prototype pollution, fix stream draining in `WorkbookReader`, replace the vulnerable `uuid` runtime dependency with an internal RFC4122 v4 helper in `lib/utils/utils.js` (using `crypto.randomUUID` when available with `crypto.randomBytes`/math fallback), and update `index.d.ts` + `spec/typescript/exceljs.spec.ts`.

**Tech Stack:** Node.js (>=18.0.0), CommonJS, Mocha + Chai, ESLint, Grunt + Babel + Browserify + Terser, TypeScript declarations (`index.d.ts`).

**Spec:** `docs/superpowers/specs/2026-11-16-exceljs-modernization-design.md`

## Global Constraints

- `engines.node` in `package.json`: `">=18.0.0"`
- Zero runtime vulnerabilities in `npm audit --omit=dev`
- Zero changes to existing public API behavior or XML output structure
- All unit tests, integration tests, ESLint checks, and TypeScript tests must pass

## Review Focus

- `_.deepMerge` or `Cell.note` receiving JSON with `{"__proto__": {"polluted": true}}` or `{"constructor": {"prototype": {"polluted": true}}}` must not pollute `Object.prototype`.
- `DefinedNames`, `CellMatrix`, and `DataValidations` receiving `"__proto__"` or `"constructor"` as a name, sheetName, or address must not pollute or read from `Object.prototype`.
- `WorksheetXform` reconciling `x14:cfRule` with `id="__proto__"` or `ref="__proto__"` must not pollute `Object.prototype`.
- `WorkbookReader` reading an `.xlsx` stream where `xl/worksheets/sheet1.xml` precedes `xl/sharedStrings.xml` and/or `xl/workbook.xml` must not truncate the deferred temp file and must resolve shared strings and worksheet metadata properly.
- `WorksheetWriter.findCell('B2')` and `WorksheetWriter.findCell(2, 2)` must return the cell at column 2 (previously returned `undefined` because of `address.column` typo).

---

### Task 1: Prototype Pollution Hardening

**Files:**
- Modify: `lib/utils/under-dash.js`
- Modify: `lib/doc/defined-names.js`
- Modify: `lib/utils/cell-matrix.js`
- Modify: `lib/doc/data-validations.js`
- Modify: `lib/xlsx/xform/sheet/worksheet-xform.js`
- Test: `spec/unit/utils/under-dash.spec.js`
- Test: `spec/unit/doc/defined-names.spec.js`

**Interfaces:**
- Consumes: `lib/utils/under-dash.js` (`_.deepMerge`, `_.keyBy`, `_.isEqual`), `lib/doc/note.js` (`Note#model`)
- Produces: Prototype-pollution-safe `_.deepMerge`, `_.keyBy`, `DefinedNames`, `CellMatrix`, `DataValidations`, and `mergeConditionalFormattings`

- [ ] **Step 1: Write the failing prototype pollution tests in `spec/unit/utils/under-dash.spec.js` and `spec/unit/doc/defined-names.spec.js`**

Add to `spec/unit/utils/under-dash.spec.js`:
```js
  describe('deepMerge prototype pollution', () => {
    afterEach(() => {
      delete Object.prototype.polluted;
    });

    it('ignores __proto__, constructor, and prototype keys in deepMerge', () => {
      const malicious = JSON.parse('{"__proto__": {"polluted": "yes"}, "constructor": {"prototype": {"polluted": "yes"}}, "safe": 1}');
      const target = _.deepMerge({}, malicious);
      expect({}.polluted).to.equal(undefined);
      expect(target.safe).to.equal(1);
    });

    it('ignores __proto__ in keyBy', () => {
      const items = [{id: '__proto__', val: 1}, {id: 'a', val: 2}];
      const keyed = _.keyBy(items, 'id');
      expect({}.val).to.equal(undefined);
      expect(keyed.a).to.deep.equal({id: 'a', val: 2});
    });
  });
```

Add to `spec/unit/doc/defined-names.spec.js`:
```js
  it('does not pollute Object.prototype with __proto__ definedName or sheetName', () => {
    const dn = new DefinedNames();
    dn.model = [
      {name: '__proto__', ranges: ['Sheet1!$A$1']},
      {name: 'valid', ranges: ['__proto__!$A$1']},
    ];
    expect({}.sheets).to.equal(undefined);
    expect(dn.getMatrix('__proto__')).to.be.an('object');
  });
```

- [ ] **Step 2: Run unit tests to verify the new test fails**

Run: `./node_modules/.bin/mocha --require spec/config/setup --require spec/config/setup-unit spec/unit/utils/under-dash.spec.js spec/unit/doc/defined-names.spec.js`
Expected: FAIL on `ignores __proto__, constructor, and prototype keys in deepMerge` (`expect({}.polluted).to.equal(undefined)` fails with `'yes'`)

- [ ] **Step 3: Guard against unsafe keys in `lib/utils/under-dash.js`, `lib/doc/defined-names.js`, `lib/utils/cell-matrix.js`, `lib/doc/data-validations.js`, and `lib/xlsx/xform/sheet/worksheet-xform.js`**

In `lib/utils/under-dash.js`:
- Add `const UNSAFE_KEYS = new Set(['__proto__', 'prototype', 'constructor']);`
- In `keyBy(a, p)`: skip or use `Object.create(null)` / guard `if (!UNSAFE_KEYS.has(v[p])) o[v[p]] = v;`
- In `isEqual`: replace `b.hasOwnProperty(key)` with `Object.prototype.hasOwnProperty.call(b, key)`
- In `deepMerge`: in `assignValue(val, key)`, return immediately if `UNSAFE_KEYS.has(key)`. Use `Object.prototype.hasOwnProperty.call(target, key) ? target[key] : undefined` when reading `src`.

In `lib/doc/defined-names.js`:
- Initialize `this.matrixMap = Object.create(null);` in `constructor()` and `set model(value)`.

In `lib/utils/cell-matrix.js`:
- Initialize `this.sheets = Object.create(null);` in `constructor()`.

In `lib/doc/data-validations.js`:
- Guard `add`, `find`, `remove` or ignore keys in `['__proto__', 'prototype', 'constructor']`.

In `lib/xlsx/xform/sheet/worksheet-xform.js`:
- In `mergeRule(rule, extRule)`: skip keys in `['__proto__', 'prototype', 'constructor']`.
- In `mergeConditionalFormattings`: create `cfMap = Object.create(null)` and `ruleMap = Object.create(null)`.
- In `reconcile`: use `Object.create(null)` for `rels`, `options.commentsMap`, and `options.hyperlinkMap`.

- [ ] **Step 4: Run unit tests to verify they pass**

Run: `./node_modules/.bin/mocha --require spec/config/setup --require spec/config/setup-unit spec/unit/utils/under-dash.spec.js spec/unit/doc/defined-names.spec.js`
Expected: PASS

---

### Task 2: Fix `WorkbookReader` Streaming Bugs & Runtime Typos

**Files:**
- Modify: `lib/stream/xlsx/workbook-reader.js`
- Modify: `lib/stream/xlsx/worksheet-writer.js`
- Modify: `lib/utils/utils.js`
- Modify: `lib/doc/table.js`
- Modify: `lib/utils/xml-stream.js`
- Test: `spec/unit/utils/utils.spec.js`
- Test: `spec/unit/doc/worksheet-table.spec.js`
- Test: `spec/unit/doc/worksheet-writer.spec.js`

**Interfaces:**
- Consumes: `WorkbookReader#parse`, `WorkbookReader#_parseWorksheet`, `WorksheetWriter#findCell`, `utils.toIsoDateString`, `Table#displayName`
- Produces: Working `WorkbookReader` for deferred worksheets on Node 18/20/22, fixed `toIsoDateString(dt: Date) -> string`, fixed `Table#displayName` getter/setter, fixed `WorksheetWriter#findCell(r, c)`

- [ ] **Step 1: Write failing unit tests for `utils.toIsoDateString`, `Table#displayName`, and `WorksheetWriter#findCell`**

Add to `spec/unit/utils/utils.spec.js`:
```js
  describe('toIsoDateString', () => {
    it('formats a Date as YYYY-MM-DD', () => {
      const dt = new Date(Date.UTC(2023, 4, 9, 12, 30, 0));
      expect(utils.toIsoDateString(dt)).to.equal('2023-05-09');
    });
  });
```

Add to `spec/unit/doc/worksheet-table.spec.js`:
```js
    it('gets and sets displayName', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('blort');
      const table = addTable('A1', ws);
      expect(table.displayName).to.equal('TestTable');
      table.displayName = 'CustomDisplayName';
      expect(table.displayName).to.equal('CustomDisplayName');
    });
```

Add to `spec/unit/doc/worksheet-writer.spec.js`:
```js
  it('finds a cell in uncommitted row via findCell', () => {
    const mockWorkbook = {
      _openStream() {
        return new StreamBuf();
      },
      styles: {addStyleModel: () => 0},
    };
    const writer = new WorksheetWriter({id: 1, workbook: mockWorkbook});
    writer.getCell('B2').value = 'hello';
    expect(writer.findCell('B2').value).to.equal('hello');
    expect(writer.findCell(2, 2).value).to.equal('hello');
  });
```

- [ ] **Step 2: Run the unit tests and the two failing integration tests to confirm failure**

Run: `./node_modules/.bin/mocha --require spec/config/setup --require spec/config/setup-unit spec/unit/utils/utils.spec.js spec/unit/doc/worksheet-table.spec.js spec/unit/doc/worksheet-writer.spec.js`
Expected: FAIL (`dt.toIsoString is not a function`, `expected 'TestTable' to equal 'CustomDisplayName'`, `Cannot read properties of undefined (reading 'value')`)

- [ ] **Step 3: Implement the fixes**

1. In `lib/stream/xlsx/workbook-reader.js`:
   - Track whether an entry was already piped to a temp file (`let drained = false;` or skip `entry.autodrain()` when piped to `tempStream`). Specifically, in `switch (entry.path)`, when piping `entry` to `tempStream`, do NOT call `entry.autodrain()` at the bottom of the loop because `entry.pipe(tempStream)` already consumes the entry stream (and calling `entry.autodrain()` before `tempStream` finishes steals the data in `parse()` when `if (this.sharedStrings && this.workbookRels && this.model)` is false).
   - Also require `this.model` in the condition on line 110: `if (this.sharedStrings && this.workbookRels && this.model)` so worksheets appearing before `xl/workbook.xml` are also deferred until `this.model` is parsed.
   - In `_parseWorksheet(iterator, sheetNo)` line 303: guard `this.model` with `const matchingSheet = matchingRel && this.model && (this.model.sheets || []).find(sheet => sheet.rId === matchingRel.Id);`.
   - Remove the obsolete `if (!fileStream[Symbol.asyncIterator])` Node v8 check on lines 142-146.
   - In `WorkbookWriter` (`lib/stream/xlsx/workbook-writer.js` line 20): change `require('../../xlsx/xml/theme1.js')` to `require('../../xlsx/xml/theme1')` to fix ESLint `import/extensions`.

2. In `lib/utils/utils.js`:
   - Fix `toIsoDateString(dt)` on line 165: `return dt.toISOString().substr(0, 10);`.

3. In `lib/doc/table.js`:
   - Fix `get displayName()` on line 408: `return this.table.displayName || this.table.name;`.
   - Add `set displayName(value) { this.table.displayName = value; }` and keep `set displayNamename(value)` delegating to it for backward compatibility.

4. In `lib/stream/xlsx/worksheet-writer.js`:
   - Fix `findCell(r, c)` on line 421: `return row ? row.findCell(address.col) : undefined;` (was `address.column`).

5. In `lib/utils/xml-stream.js`:
   - Line 22: change `xml.push(tmp.join(""));` to `xml.push(tmp.join(''));` to fix ESLint `quotes`.

- [ ] **Step 4: Run unit tests, integration tests (`issue-1328` & `test-pr-1431`), and ESLint on `lib/`**

Run: `./node_modules/.bin/mocha --require spec/config/setup --require spec/config/setup-unit spec/unit/utils/utils.spec.js spec/unit/doc/worksheet-table.spec.js spec/unit/doc/worksheet-writer.spec.js && ./node_modules/.bin/mocha --require spec/config/setup spec/integration/issues/issue-1328-xlsx-worksheet-reader-date.spec.js spec/integration/pr/test-pr-1431.spec.js && ./node_modules/.bin/eslint lib`
Expected: PASS with 0 ESLint errors

---

### Task 3: Dependency Updates, Zero-Vuln UUID, and Obsolete Code Cleanup

**Files:**
- Modify: `package.json`
- Modify: `excel.js`
- Modify: `lib/utils/utils.js`
- Modify: `lib/xlsx/xform/sheet/cf-ext/cf-rule-ext-xform.js`
- Modify: `lib/utils/iterate-stream.js`
- Modify: `lib/utils/parse-sax.js`
- Modify: `lib/xlsx/xlsx.js`
- Modify: `lib/doc/row.js`
- Modify: `lib/doc/column.js`
- Modify: `lib/utils/zip-stream.js`
- Modify: `lib/xlsx/xform/strings/phonetic-text-xform.js`
- Modify: `lib/xlsx/xform/strings/shared-string-xform.js`
- Test: `spec/unit/utils/utils.spec.js`

**Interfaces:**
- Consumes: `crypto` built-in module
- Produces: `utils.uuidv4() -> string` (RFC4122 v4 UUID), updated `package.json` with `engines.node >= 18.0.0` and 0 runtime security advisories

- [ ] **Step 1: Write unit test for `utils.uuidv4()` in `spec/unit/utils/utils.spec.js`**

Add to `spec/unit/utils/utils.spec.js`:
```js
  describe('uuidv4', () => {
    it('generates valid RFC4122 v4 UUIDs', () => {
      const id1 = utils.uuidv4();
      const id2 = utils.uuidv4();
      expect(id1).to.match(/^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i);
      expect(id1).to.not.equal(id2);
    });
  });
```

- [ ] **Step 2: Run unit test to verify it fails**

Run: `./node_modules/.bin/mocha --require spec/config/setup --require spec/config/setup-unit spec/unit/utils/utils.spec.js`
Expected: FAIL (`utils.uuidv4 is not a function`)

- [ ] **Step 3: Implement `utils.uuidv4`, replace `uuid` dependency, clean up Node 8/10 shims, and update `package.json`**

1. In `lib/utils/utils.js`:
   - Require `crypto`: `const crypto = require('crypto');`
   - Add `uuidv4()` to `utils`:
     ```js
     uuidv4() {
       if (typeof crypto.randomUUID === 'function') {
         return crypto.randomUUID();
       }
       const rnds = crypto.randomBytes(16);
       rnds[6] = (rnds[6] & 0x0f) | 0x40;
       rnds[8] = (rnds[8] & 0x3f) | 0x80;
       const hex = rnds.toString('hex');
       return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20, 32)}`;
     },
     ```
   - Simplify `objectFromProps(props, value = null)` using `Object.fromEntries(props.map(property => [property, value]))` or `reduce`.

2. In `lib/xlsx/xform/sheet/cf-ext/cf-rule-ext-xform.js`:
   - Replace `const {v4: uuidv4} = require('uuid');` with `const {uuidv4} = require('../../../../utils/utils');`.

3. Remove obsolete Node 8/10 workarounds:
   - `lib/utils/iterate-stream.js`: remove custom `once` TODO comment; keep clean stream iteration.
   - `lib/utils/parse-sax.js`: remove `// TODO: Remove once node v8 is deprecated` comment.
   - `lib/xlsx/xlsx.js`: remove `// TODO: Remove once node v8 is deprecated` comment.
   - Replace direct `.hasOwnProperty(...)` calls in `lib/doc/row.js:184`, `lib/doc/column.js:195`, `lib/utils/zip-stream.js:26`, `lib/xlsx/xform/strings/phonetic-text-xform.js:29`, `lib/xlsx/xform/strings/shared-string-xform.js:33` with `Object.prototype.hasOwnProperty.call(...)`.

4. In `package.json`:
   - Update `"engines": { "node": ">=18.0.0" }`.
   - Remove `"uuid"` from `"dependencies"` (since it was only used in `cf-rule-ext-xform.js` and caused the moderate CVE advisory).
   - Update safe CJS-compatible runtime dependencies:
     - `"archiver": "^7.0.1"`
     - `"dayjs": "^1.11.13"`
     - `"jszip": "^3.10.1"`
     - `"saxes": "^6.0.0"`
     - `"tmp": "^0.2.3"`
   - Run `npm install` (with network enabled) and verify `npm audit --omit=dev` reports **0 vulnerabilities**.

- [ ] **Step 4: Run `npm audit --omit=dev` and unit + integration tests**

Run: `npm audit --omit=dev && ./node_modules/.bin/mocha --require spec/config/setup --require spec/config/setup-unit spec/unit --recursive`
Expected: `found 0 vulnerabilities` and all unit tests PASS

---

### Task 4: TypeScript Definitions (`index.d.ts`) & TypeScript Test Suite

**Files:**
- Modify: `index.d.ts`
- Modify: `spec/typescript/exceljs.spec.ts`

**Interfaces:**
- Consumes: Runtime classes `Workbook`, `Worksheet`, `Row`, `Column`, `Table`, `ModelContainer`, `WorksheetReader`, `makePivotTable`
- Produces: Accurate `index.d.ts` definitions and passing `npm run test:typescript`

- [ ] **Step 1: Update `spec/typescript/exceljs.spec.ts` to test modern streaming and new TS types (`getTables`, `addPivotTable`, `Row.dimensions`, `Column.values`)**

Update `spec/typescript/exceljs.spec.ts`:
```ts
import 'regenerator-runtime/runtime';
import { PassThrough } from 'stream';
import { expect } from 'chai';
import ExcelJS from '../../index';

describe('typescript', () => {
  it('can create and buffer xlsx', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');
    ws.getCell('A1').value = 7;
    const buffer = await wb.xlsx.writeBuffer({
      useStyles: true,
      useSharedStrings: true,
    });

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(buffer);
    const ws2 = wb2.getWorksheet('blort');
    expect(ws2.getCell('A1').value).to.equal(7);
  });

  it('can create and stream xlsx', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');
    ws.getCell('A1').value = 7;

    const stream = new PassThrough();
    await wb.xlsx.write(stream);
    stream.end();

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.read(stream);
    const ws2 = wb2.getWorksheet('blort');
    expect(ws2.getCell('A1').value).to.equal(7);
  });

  it('types getTables, row dimensions, column values, and addPivotTable accurately', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.addRows([
      ['A', 'B', 'C'],
      ['a1', 'b1', 10],
    ]);

    const rowDims: { min: number; max: number } | null = ws.getRow(1).dimensions;
    expect(rowDims).to.deep.equal({ min: 1, max: 3 });

    const colVals: readonly ExcelJS.CellValue[] = ws.getColumn(1).values;
    expect(colVals.length).to.be.greaterThan(0);

    const tables: ExcelJS.Table[] = ws.getTables();
    expect(tables).to.be.an('array');

    const ws2 = wb.addWorksheet('Pivot');
    const pt = ws2.addPivotTable({
      sourceSheet: ws,
      rows: ['A'],
      columns: ['B'],
      values: ['C'],
      metric: 'sum',
    });
    expect(pt.metric).to.equal('sum');
  });
});
```

- [ ] **Step 2: Update `index.d.ts` with missing/corrected types**

In `index.d.ts`:
- Change `Row.dimensions` from `readonly dimensions: number;` to `readonly dimensions: { min: number; max: number } | null;`.
- Change `Column.values` from `values: CellValue;` to `values: ReadonlyArray<CellValue>;`.
- Change `Worksheet.getTables()` from `getTables(): [Table, void][];` to `getTables(): Table[];`.
- Add `PivotTableModel` and `PivotTable` interfaces and `addPivotTable(model: PivotTableModel): PivotTable;` on `Worksheet` (and `pivotTables: PivotTable[];` on `Workbook` and `Worksheet`):
  ```ts
  export interface PivotTableModel {
  	sourceSheet: Worksheet;
  	rows: string[];
  	columns: string[];
  	values: string[];
  	metric?: 'sum';
  }

  export interface PivotTableCacheField {
  	name: string;
  	sharedItems: any[] | null;
  }

  export interface PivotTable {
  	sourceSheet: Worksheet;
  	rows: number[];
  	columns: number[];
  	values: number[];
  	metric: 'sum';
  	cacheFields: PivotTableCacheField[];
  	cacheId: string;
  }
  ```
- Export `class ModelContainer`:
  ```ts
  export class ModelContainer {
  	model: any;
  	constructor(model: any);
  	readonly xlsx: Xlsx;
  }
  ```
- In `stream.xlsx.WorksheetReader`:
  - Fix `readonly dimensions: Range;` (was `dimensions(): number;`)
  - Fix `readonly columns: Column[];` (was `columns(): number;`)
  - Fix `getColumn(c: number | string): Column;`
  - Add `id: number | string;`, `name: string;`, `state?: WorksheetState;`

- [ ] **Step 3: Build `dist/` and run `npm run test:typescript`**

Run: `npm run build && npm run test:typescript`
Expected: PASS (3 passing TypeScript tests and `tsc --noEmit index.d.ts` exits 0)

---

### Task 5: Full Verification (Build, Unit, Integration, End-to-End, TypeScript, Lint)

**Files:**
- Modify: none (verification only)

**Interfaces:**
- Consumes: All tasks 1–4
- Produces: Verified build artifacts in `dist/` and green test suites

- [ ] **Step 1: Run full build, lint, unit, integration, end-to-end, and TypeScript test suites**

Run:
1. `./node_modules/.bin/eslint lib`
2. `npm run build`
3. `npm run test:unit`
4. `npm run test:integration` (with sandbox bypass for the Express local port test)
5. `npm run test:end-to-end`
6. `npm run test:typescript`
7. `npm audit --omit=dev`

Expected: All suites pass with 0 errors and 0 runtime vulnerabilities.
