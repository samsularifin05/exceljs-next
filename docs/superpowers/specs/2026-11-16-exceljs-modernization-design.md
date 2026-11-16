# ExcelJS Modernization & Security Hardening Design

## Intent & Why
ExcelJS (`4.4.0`) was last updated ~3 years ago. The goal is to fix security vulnerabilities (both in source code and dependencies), fix broken streaming/runtime bugs on modern Node.js, clean up redundant/obsolete code, and bring TypeScript definitions (`index.d.ts`) into full sync with the runtime API — **without changing existing business logic or breaking public CJS/browser APIs**.

## Goals
1. **Security Hardening:**
   - Eliminate prototype pollution vectors in `lib/utils/under-dash.js` (`deepMerge`, `keyBy`), `lib/doc/defined-names.js`, `lib/utils/cell-matrix.js`, `lib/doc/data-validations.js`, and `lib/xlsx/xform/sheet/worksheet-xform.js` (`mergeConditionalFormattings`, `mergeRule`, `reconcile`).
   - Achieve 0 runtime dependency vulnerabilities (`npm audit --omit=dev`).
2. **Bug Fixes (Zero Logic Change to Working Features):**
   - Fix `WorkbookReader` streaming worksheet deferral bug on Node 18/20/22 where `entry.autodrain()` drains the zip entry stream concurrently with `entry.pipe(tempStream)`, truncating deferred worksheets when `sharedStrings.xml` appears after `sheetN.xml` in the archive (fixes integration tests `issue-1328` and `test-pr-1431`).
   - Guard `WorkbookReader._parseWorksheet` when `this.model` is not yet populated before worksheet parsing.
   - Fix clear typos/bugs in existing methods:
     - `lib/utils/utils.js`: `toIsoDateString` (`dt.toIsoString().subsstr(0, 10)` -> `dt.toISOString().substr(0, 10)`).
     - `lib/doc/table.js`: `get displayName()` typo (`this.table.displyName` -> `this.table.displayName`) and `set displayNamename(value)` -> `set displayName(value)` (keeping `displayNamename` as an alias for backward compatibility).
     - `lib/stream/xlsx/worksheet-writer.js`: `findCell` uses `address.column` instead of `address.col` (returned by `colCache.getAddress`).
     - Fix ESLint errors in `lib/stream/xlsx/workbook-writer.js` and `lib/utils/xml-stream.js`.
3. **Dependency & Engine Updates:**
   - Update `engines.node` to `>=18.0.0`.
   - Upgrade safe runtime dependencies compatible with CommonJS and Browserify: `archiver` (`^7.0.1`), `dayjs` (`^1.11.23`), `jszip` (`^3.10.1`), `saxes` (`^6.0.0`), `tmp` (`^0.2.5`), and update `uuid` to `^10.0.0` (last CJS-supporting release before v11/v14 ESM-only) or safe CJS-compatible version so `npm audit --omit=dev` has 0 vulnerabilities, or use a tiny RFC4122 v4 helper in `lib/utils/utils.js` using `crypto` / `Math.random` fallback for browser compatibility.
4. **Code Cleanup & De-duplication:**
   - Remove obsolete Node v8/v10 stream/events workarounds (`iterate-stream.js`, `parse-sax.js`, `workbook-reader.js`, `xlsx.js`, `excel.js`, `utils.js`).
   - Consolidate duplicated worksheet protection hashing logic between `lib/doc/worksheet.js` and `lib/stream/xlsx/worksheet-writer.js`, and duplicated `hasOwnProperty` calls using `Object.prototype.hasOwnProperty.call`.
5. **TypeScript Support (`index.d.ts` & `spec/typescript`):**
   - Fix inaccurate types in `index.d.ts`:
     - `Worksheet.getTables(): Table[]` (was `[Table, void][]`).
     - `Worksheet.addPivotTable(model: PivotTableModel): PivotTable` and `PivotTable` / `PivotTableModel` interfaces.
     - `Row.dimensions: { min: number; max: number } | null` (was `number`).
     - `Column.values: ReadonlyArray<CellValue>` (was `CellValue`).
     - `stream.xlsx.WorksheetReader`: `dimensions: Range` and `columns: Column[]` (were methods returning `number`), `getColumn(c: number | string): Column`.
     - Export `ModelContainer` class declaration.
   - Fix `spec/typescript/exceljs.spec.ts` to use `wb2.xlsx.read(stream)` instead of removed `wb2.xlsx.createInputStream()` so `npm run test:typescript` passes.

## Non-Goals & Anti-Goals
- **Non-Goal:** Rewriting the 171 JS source files into TypeScript (high regression risk).
- **Non-Goal:** Breaking CommonJS `require('exceljs')` or Browserify `dist/` bundles.
- **Anti-Goal:** Changing any existing public method behavior, XML output structure, or cell value coercion rules.

## Approaches Considered
1. **Option A — Conservative Modernization (Chosen):** Keep CommonJS architecture and proven stream libraries (`readable-stream@3`, `unzipper@0.10`, `fast-csv@4`), fix all security advisories in code + runtime deps, fix Node 22 streaming bugs, clean up dead Node 8/10 shims, and update `index.d.ts` + TS tests.
2. **Option B — Full Major Dependency & ESM Overhaul:** Upgrade `archiver@8`, `uuid@14`, `readable-stream@4`, `fast-csv@5`, `unzipper@0.12`. Rejected because `archiver@8` and `uuid@14` are ESM-only (breaking CJS `require` and Browserify), and `unzipper@0.12` / `readable-stream@4` alter stream timing in ways that break existing consumers.
