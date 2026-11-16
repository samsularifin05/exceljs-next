# Changelog — `exceljs-next`

> **Note:** [samsularifin05/exceljs-next](https://github.com/samsularifin05/exceljs-next) is a maintained and security-hardened fork of the original [exceljs/exceljs](https://github.com/exceljs/exceljs) (which has been inactive since `v4.4.0` in January 2024). It retains 100% backward compatibility with the original API while fixing critical security vulnerabilities, Node.js 16–22+ streaming bugs, Excel file corruption issues, TypeScript & ESM support, and adding long-requested features.

---

## [4.5.0] - 2026-11-17

### Security Fixes (0 Vulnerabilities)
- **Prototype Pollution Hardening:** Blocked `__proto__`, `constructor`, and `prototype` keys across `_.deepMerge` (including `cell.note` serialization), `_.keyBy`, `DefinedNames`, `CellMatrix`, `DataValidations`, `WorkSheetXform` (`x14:cfRule` / `mergeConditionalFormattings`), `ws.toJSON()`, `ws.fromJSON()`, and `renderTemplate()`.
- **Anti-DoS / Heap OOM Protection:** Added row/column bounds validation (`1..1048576` rows, `1..16384` columns) in `colCache.decodeAddress` and rejected oversized `<mergeCell>` ranges (`> 1,000,000` cells) to prevent uncatchable heap OOM crashes on crafted `.xlsx` files.
- **CSV Prototype Lookup Protection:** Restricted `SpecialValues` lookup in `lib/csv/csv.js` using `Object.prototype.hasOwnProperty.call` so cells containing `"constructor"`, `"toString"`, or `"__proto__"` remain strings.
- **Zero Dependency Vulnerabilities & Smaller Footprint:** Replaced `uuid` with built-in `crypto.randomUUID` / `crypto.randomBytes` (`utils.uuidv4()`), replaced `tmp` with built-in `fs`/`os`/`crypto` (`createTempFile()` with exclusive `'wx'` flags and `0o600` permissions), updated runtime dependencies (`saxes@^6.0.0`, `dayjs@^1.11.13`), and added `overrides` so `npm audit` reports **0 vulnerabilities**.

### New Features
1. **File Encryption & Decryption — Password to Open (`.xlsx` `MS-OFFCRYPTO`)**
   - Added pure-JS MS-CFB (OLE2) + **MS-OFFCRYPTO** support (`lib/utils/offcrypto.js`) wired into `workbook.xlsx.writeFile()`, `writeBuffer()`, `write()`, `readFile()`, `read()`, and `load()`.
   - Encrypts `.xlsx` files with **Agile Encryption (AES-256-CBC + SHA-512 + HMAC-SHA512)** when `{ password }` is provided.
   - Decrypts both **Agile (`4.4`)** and **Standard (`3.2`/`4.2`)** encrypted `.xlsx` files and validates incorrect/missing passwords.
2. **Native Excel Charts (`worksheet.addChart` & `worksheet.getCharts`)**
   - Added native OOXML chart generation (`ChartXform` + `TwoCellAnchorXform` `<xdr:graphicFrame>`) supporting `'column'`, `'bar'`, `'line'`, `'area'`, `'pie'`, and `'doughnut'` charts.
   - Supports multiple series, custom hex colors, `valAxis` (`numFmt`, `majorUnit`), `legend` positioning, and `dataLabels` (`showVal`, `showPercent`, `showCatName`, `showLeaderLines`, `separator` — enabled by default for `pie` and `doughnut` charts).
3. **Template Engine (`workbook.renderTemplate` & `worksheet.renderTemplate`)**
   - Added `lib/utils/template-engine.js` to replace `{{placeholder}}` and nested dot-path tokens (`{{customer.name}}`), including RichText runs.
   - Preserves native JS/Excel types (`number`, `Date`, `boolean`) when a cell contains a single exact placeholder.
   - Automatically duplicates and populates array loop rows (`{{items.name}}`, `{{items.$index}}`, `{{items.$number}}`) while preserving template cell styles, or removes empty loop rows (`removeEmptyLoopRows`).
4. **JSON Helpers (`worksheet.fromJSON` & `worksheet.toJSON`)**
   - Added `ws.fromJSON(records, { headers, headerLabels, includeHeader, headerStyle, autoFit })` to populate worksheets directly from arrays of objects.
   - Added `ws.toJSON({ headerRow, startRow, endRow, includeEmpty, raw })` to export worksheets back to clean JSON objects (unwrapping formulas, rich text, and hyperlinks).
5. **Auto-Fit Column Widths & Row Heights**
   - Added `worksheet.autoFitColumns({ minWidth = 9, maxWidth = 100, padding = 2, onlyEmptyWidth = false })`.
   - Added `worksheet.autoFitRows({ minHeight = 15, maxHeight = 409, padding = 4, onlyEmptyHeight = false })` supporting `\n` newlines, `wrapText` line wrapping estimation, and font sizes.
6. **Enhanced Pivot Tables (`worksheet.addPivotTable`)**
   - Support **multiple Pivot Tables** within the same workbook (`pivotTable1.xml`, `pivotTable2.xml`, etc., with distinct `cacheId`s).
   - Support **11 aggregation metrics (`metric`)**: `'sum'`, `'count'`, `'average'`, `'max'`, `'min'`, `'product'`, `'countNums'`, `'stdDev'`, `'stdDevP'`, `'var'`, `'varP'`.
   - Support **multiple `values` fields** and **`sourceTable`** (in addition to `sourceSheet`).
   - Escape XML special characters (`&`, `<`, `>`, `"`, `'`) and handle `null`/`undefined` values (`<m />`) in Pivot Table cache definitions and records.
7. **Streaming Writer & Reader Parity (`ExcelJS.stream.xlsx`)**
   - Added `wsWriter.addTable()`, `getTable()`, `removeTable()`, and `getTables()` to `WorksheetWriter` & `WorkbookWriter`.
   - Added `wsWriter.addImage(imageId, range)` and `getImages()` to `WorksheetWriter` & `WorkbookWriter`.
   - Added `<extLst>` rendering for extended conditional formatting (`dataBar`, custom `iconSet`) in `WorksheetWriter`.
   - Added `<mergeCell>` parsing to `WorksheetReader` (`sheetReader.mergeCells` array and `'mergeCell'` event).
8. **Column Page Breaks (`colBreaks`)**
   - Added `column.addPageBreak(top, bottom)` and `worksheet.colBreaks` (in-memory & streaming), plus full `<rowBreaks>` and `<colBreaks>` XML parsing on load.
9. **Workbook Structure Protection (`workbook.protect` & `workbook.unprotect`)**
   - Added `await workbook.protect(password, options)` (`lockStructure`, `lockWindows`, `lockRevision`, `spinCount` with SHA-512 hashing) and `workbook.unprotect()`, including `<workbookProtection>` serialization and parsing.
10. **Native ESM, Node.js `>=16.0.0`, `pnpm`, & Strict TypeScript Support**
    - Added `excel.mjs` (ESM entry point), `index.d.mts` (ESM TypeScript declarations), and separated `"import"` / `"require"` `"exports"` in `package.json`.
    - Lowered `"engines.node"` to `">=16.0.0"` (supporting Node 16 through Node 22+).
    - Migrated scripts to `pnpm` (`packageManager: "pnpm@12.6.0"`, `pnpm.onlyBuiltDependencies`, `pnpm.overrides`) and added local Chrome auto-detection in `Gruntfile.js`.
11. **Documentation & Interactive Demo**
    - Added interactive browser playground [demo.html](demo.html).
    - Added full 3-language documentation: [README.md](README.md) (English), [README_id.md](README_id.md) (Indonesian), and [README_zh.md](README_zh.md) (Chinese).
    - Added 12-page GitHub Wiki tutorials in `wiki/`.

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
- **TypeScript (`index.d.ts`, `index.d.mts`, & `tsconfig.json`):**
  - Updated types for `PivotTable`, `PivotTableModel`, `PivotTableMetric`, `WorkbookProtection`, `AutoFitColumnsOptions`, `AutoFitRowsOptions`, `ToJSONOptions`, `FromJSONOptions`, `TemplateRenderOptions`, `ChartOptions`, `ChartDataLabelsOptions`, `XlsxReadOptions.password`, `XlsxWriteOptions.password`, `ColBreak`, `Column.addPageBreak`, `Column.values`, `Row.dimensions`, `Worksheet.getTables(): Table[]`, `WorksheetReader` (`EventEmitter` strict-mode fix), and `ModelContainer`.
