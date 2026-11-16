# Changelog (Maintained Fork)

> **Note:** This repository is a maintained and security-hardened fork of the original [exceljs/exceljs](https://github.com/exceljs/exceljs) (which has been inactive since v4.4.0 in January 2024). It retains 100% backward compatibility with the original API while fixing critical security vulnerabilities, Node.js 18/20/22+ streaming bugs, Excel file corruption issues, TypeScript definitions, and adding long-requested features.

---

## [4.5.0] - 2026-11-16

### Security Fixes (0 Vulnerabilities)
- **Prototype Pollution Hardening:** Blocked `__proto__`, `constructor`, and `prototype` keys across `_.deepMerge` (including `cell.note` serialization), `_.keyBy`, `DefinedNames`, `CellMatrix`, `DataValidations`, and `WorkSheetXform` (`x14:cfRule` / `mergeConditionalFormattings`).
- **Anti-DoS / Heap OOM Protection:** Added row/column bounds validation (`1..1048576` rows, `1..16384` columns) in `colCache.decodeAddress` and rejected oversized `<mergeCell>` ranges (`> 1,000,000` cells) to prevent uncatchable heap OOM crashes on crafted `.xlsx` files.
- **CSV Prototype Lookup Protection:** Restricted `SpecialValues` lookup in `lib/csv/csv.js` using `Object.prototype.hasOwnProperty.call` so cells containing `"constructor"`, `"toString"`, or `"__proto__"` remain strings.
- **Zero Dependency Vulnerabilities:** Replaced the vulnerable `uuid` package with built-in `crypto.randomUUID` / `crypto.randomBytes` (`utils.uuidv4()`), updated runtime dependencies (`saxes@^6.0.0`, `dayjs@^1.11.13`, `tmp@^0.2.7`), and added `overrides` so `npm audit` reports **0 vulnerabilities**.

### New Features
1. **Enhanced Pivot Tables (`worksheet.addPivotTable`)**
   - Support **multiple Pivot Tables** within the same workbook (`pivotTable1.xml`, `pivotTable2.xml`, etc., with distinct `cacheId`s).
   - Support **11 aggregation metrics (`metric`)**: `'sum'`, `'count'`, `'average'`, `'max'`, `'min'`, `'product'`, `'countNums'`, `'stdDev'`, `'stdDevP'`, `'var'`, `'varP'`.
   - Support **multiple `values` fields** in a single Pivot Table.
   - Support **`sourceTable`** (creating a Pivot Table directly from a `Table` instance in addition to `sourceSheet`).
   - Escape XML special characters (`&`, `<`, `>`, `"`, `'`) and handle `null`/`undefined` values (`<m />`) in Pivot Table cache definitions and records.
2. **Streaming Writer & Reader Parity (`ExcelJS.stream.xlsx`)**
   - Added `wsWriter.addTable()`, `getTable()`, `removeTable()`, and `getTables()` to `WorksheetWriter` & `WorkbookWriter`.
   - Added `wsWriter.addImage(imageId, range)` and `getImages()` to `WorksheetWriter` & `WorkbookWriter`.
   - Added `<extLst>` rendering for extended conditional formatting (`dataBar`, custom `iconSet`) in `WorksheetWriter`.
   - Added `<mergeCell>` parsing to `WorksheetReader` (`sheetReader.mergeCells` array and `'mergeCell'` event).
3. **Column Page Breaks (`colBreaks`)**
   - Added `column.addPageBreak(top, bottom)` and `worksheet.colBreaks` (in-memory & streaming), plus full `<rowBreaks>` and `<colBreaks>` XML parsing on load.
4. **Auto-Fit Column Width (`worksheet.autoFitColumns`)**
   - Added `worksheet.autoFitColumns({ minWidth = 9, maxWidth = 100, padding = 2, onlyEmptyWidth = false })` to automatically calculate column widths from cell contents, headers, rich text, formula results, dates, and font sizes.
5. **Workbook Protection (`workbook.protect` & `workbook.unprotect`)**
   - Added `await workbook.protect(password, options)` (`lockStructure`, `lockWindows`, `lockRevision`, `spinCount` with SHA-512 hashing) and `workbook.unprotect()`, including `<workbookProtection>` serialization and parsing.
6. **Native ESM Entry Point (`excel.mjs`)**
   - Added `excel.mjs` and conditional `"exports"` in `package.json` supporting both `import ExcelJS, { Workbook, ValueType } from '...'` (ESM) and `require('...')` (CommonJS).

### Bug Fixes
- **Node.js 18/20/22+ Streaming Reader (`WorkbookReader`):** Fixed `iterateStream` dropping buffered zip entries when `unzipper` emits `'end'` while paused, and prevented double-draining deferred worksheet streams (fixing unresolved `{ sharedString: index }` values).
- **UTF-8 Multi-Byte Streaming:** Used streaming `TextDecoder('utf-8', { stream: true })` per SAX parser instance so multi-byte UTF-8 characters (CJK, Arabic, Emoji) split across chunk boundaries are not corrupted.
- **Excel File Corruption Fixes:**
  - Removed illegal XML 1.0 characters `U+FFFE` and `U+FFFF` in `utils.xmlEncode()`.
  - Escaped apostrophes (`'` -> `''`) in worksheet names referenced by `Print_Area` and `Print_Titles`.
  - Assigned distinct 1024-ID VML shape blocks per worksheet (`<o:idmap data="${sheetId}"/>`) when multiple worksheets contain cell comments/notes.
  - Assigned unique `a16:creationId` per picture and fixed consecutive duplicate image placement in `WorkSheetXform`.
- **Formula & Cell Fixes:**
  - Preserved falsy formula results (`0`, `false`, `""`) in `FormulaValue._copyModel`.
  - Preserved string and boolean formula results on date-formatted cells instead of coercing them through `excelToDate`.
  - Handled shared/inline strings that mix a leading `<t>` element with `<r>` rich-text runs.
  - Fixed `<alignment wrapText="0" shrinkToFit="0"/>` being parsed as `true`.
- **Worksheet, Column & Table Fixes:**
  - Compared `Column.isCustomWidth` against `worksheet.properties.defaultColWidth` before falling back to `9`.
  - Excluded the worksheet itself (`ws !== this`) when checking for duplicate names on >31-char truncation.
  - Attached the `Worksheet` instance when loading tables (`new Table(this)`) and fixed `Table.displayName` getter/setter typos.
  - Guarded missing comment/VML metadata in `WorkSheetXform.reconcile`.
  - Fixed `utils.toIsoDateString()` and `WorksheetWriter.findCell()` typos.
- **TypeScript (`index.d.ts`):**
  - Updated types for `PivotTable`, `PivotTableModel`, `PivotTableMetric`, `WorkbookProtection`, `AutoFitColumnsOptions`, `ColBreak`, `Column.addPageBreak`, `Column.values`, `Row.dimensions`, `Worksheet.getTables(): Table[]`, `WorksheetReader`, and `ModelContainer`.
