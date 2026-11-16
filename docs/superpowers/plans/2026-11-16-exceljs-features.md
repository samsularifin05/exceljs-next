# ExcelJS 6-Feature Enhancement Suite Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement 6 new feature sets (Enhanced Pivot Tables, Streaming Parity, Column Page Breaks, Auto-Fit Columns, Workbook Protection, and Native ESM Support) with full unit/integration/TypeScript test coverage and zero regressions.

**Architecture:** Extend the existing Document (`lib/doc/`), Xform (`lib/xlsx/xform/`), Streaming (`lib/stream/xlsx/`), and TypeScript declaration (`index.d.ts`) layers modularly so existing behavior is untouched while new capabilities work seamlessly across in-memory and streaming workflows.

**Tech Stack:** Node.js (>=18.0.0), CommonJS + ESM wrapper (`excel.mjs`), Mocha + Chai, TypeScript (`index.d.ts`).

**Spec:** `docs/superpowers/specs/2026-11-16-exceljs-features-design.md`

## Global Constraints

- Preserve 100% compatibility with all existing unit, integration, end-to-end, and TypeScript tests.
- Maintain 0 ESLint errors in `lib/` and 0 vulnerabilities in `npm audit`.
- Update `index.d.ts` for every new public method, property, and option.

## Review Focus

- Adding 2+ pivot tables in the same workbook or across worksheets must generate distinct `pivotTable{n}.xml`, `pivotCacheDefinition{n}.xml`, `pivotCacheRecords{n}.xml`, and unique `cacheId` values without overwriting `pivotTable1.xml`.
- Pivot tables with `metric: 'count'` or multiple `values: ['D', 'E']` must emit valid `<dataField subtotal="..." />` and `<colFields>` entries.
- `column.addPageBreak()` must round-trip through `Workbook.xlsx.writeBuffer()` and `Workbook.xlsx.load()`.
- `worksheet.autoFitColumns()` on an empty worksheet or columns with `null`/`RichText`/`Formula`/`Date` values must not throw and must respect `minWidth` and `maxWidth`.
- `workbook.protect('secret', { lockStructure: true })` must round-trip `<workbookProtection>` attributes through `writeBuffer()` and `load()`.

---

### Task 1: Pivot Table Enhancements (Multiple Tables, Metrics, Multiple Values, `sourceTable`)

**Files:**
- Modify: `lib/doc/pivot-table.js`
- Modify: `lib/xlsx/xform/pivot-table/pivot-table-xform.js`
- Modify: `lib/xlsx/xform/pivot-table/pivot-cache-definition-xform.js`
- Modify: `lib/xlsx/xform/sheet/worksheet-xform.js`
- Modify: `lib/xlsx/xform/core/content-types-xform.js`
- Modify: `lib/xlsx/xlsx.js`
- Test: `spec/integration/workbook/pivot-tables.spec.js`

- [ ] **Step 1: Write failing tests in `spec/integration/workbook/pivot-tables.spec.js` for multiple pivot tables, metrics (`count`, `average`, `max`, `min`), multiple `values`, and `sourceTable`**
- [ ] **Step 2: Run `spec/integration/workbook/pivot-tables.spec.js` to confirm failure**
- [ ] **Step 3: Implement multi-pivot-table numbering (`tableNumber`), metric map, multi-values rendering, and `sourceTable` support**
- [ ] **Step 4: Run `spec/integration/workbook/pivot-tables.spec.js` to confirm PASS**

---

### Task 2: Streaming Writer & Reader Parity (`addTable`, `addImage`, `extLst`, `mergeCells` reader)

**Files:**
- Modify: `lib/stream/xlsx/workbook-writer.js`
- Modify: `lib/stream/xlsx/worksheet-writer.js`
- Modify: `lib/stream/xlsx/sheet-rels-writer.js`
- Modify: `lib/stream/xlsx/worksheet-reader.js`
- Test: `spec/unit/doc/worksheet-writer.spec.js`
- Test: `spec/integration/workbook-xlsx-reader.spec.js`

- [ ] **Step 1: Write failing tests for `WorksheetWriter#addTable`, `WorksheetWriter#addImage`, `WorksheetWriter` extended conditional formatting, and `WorksheetReader#mergeCells`**
- [ ] **Step 2: Run tests to confirm failure**
- [ ] **Step 3: Implement `addTable`, `addImage`, and `extLst` in `WorksheetWriter`/`WorkbookWriter`, and `<mergeCell>` tracking in `WorksheetReader`**
- [ ] **Step 4: Run tests to confirm PASS**

---

### Task 3: Column Page Breaks (`colBreaks` & `column.addPageBreak`)

**Files:**
- Create: `lib/xlsx/xform/sheet/col-breaks-xform.js`
- Modify: `lib/xlsx/xform/sheet/page-breaks-xform.js`
- Modify: `lib/doc/column.js`
- Modify: `lib/doc/worksheet.js`
- Modify: `lib/xlsx/xform/sheet/worksheet-xform.js`
- Modify: `lib/stream/xlsx/worksheet-writer.js`
- Test: `spec/unit/doc/worksheet.pagebreak.spec.js`

- [ ] **Step 1: Write failing unit and round-trip tests in `spec/unit/doc/worksheet.pagebreak.spec.js` for `column.addPageBreak()` and `<colBreaks>` XML serialization/parsing**
- [ ] **Step 2: Run `spec/unit/doc/worksheet.pagebreak.spec.js` to confirm failure**
- [ ] **Step 3: Implement `ColBreaksXform`, update `PageBreaksXform.parseOpen` to parse `id`/`max`/`min`/`man`, and wire `colBreaks` in `Column`, `Worksheet`, `WorkSheetXform`, and `WorksheetWriter`**
- [ ] **Step 4: Run `spec/unit/doc/worksheet.pagebreak.spec.js` and existing `page-breaks-xform.spec.js` to confirm PASS**

---

### Task 4: Auto-Fit Column Width (`worksheet.autoFitColumns`)

**Files:**
- Modify: `lib/doc/worksheet.js`
- Test: `spec/unit/doc/column.spec.js`

- [ ] **Step 1: Write failing unit tests in `spec/unit/doc/column.spec.js` for `worksheet.autoFitColumns(options)`**
- [ ] **Step 2: Run `spec/unit/doc/column.spec.js` to confirm failure**
- [ ] **Step 3: Implement `autoFitColumns(options)` in `lib/doc/worksheet.js`**
- [ ] **Step 4: Run `spec/unit/doc/column.spec.js` to confirm PASS**

---

### Task 5: Workbook Protection (`workbook.protect` & `workbook.unprotect`)

**Files:**
- Create: `lib/xlsx/xform/book/workbook-protection-xform.js`
- Modify: `lib/doc/workbook.js`
- Modify: `lib/xlsx/xform/book/workbook-xform.js`
- Modify: `lib/xlsx/xlsx.js`
- Test: `spec/unit/doc/workbook.spec.js`

- [ ] **Step 1: Write failing unit and round-trip tests in `spec/unit/doc/workbook.spec.js` for `workbook.protect(password, options)` and `workbook.unprotect()`**
- [ ] **Step 2: Run `spec/unit/doc/workbook.spec.js` to confirm failure**
- [ ] **Step 3: Implement `WorkbookProtectionXform` and wire `workbookProtection` in `Workbook`, `WorkbookXform`, and `XLSX`**
- [ ] **Step 4: Run `spec/unit/doc/workbook.spec.js` to confirm PASS**

---

### Task 6: Native ESM Support (`excel.mjs`, `package.json` exports) & TypeScript Definitions (`index.d.ts`)

**Files:**
- Create: `excel.mjs`
- Modify: `package.json`
- Modify: `index.d.ts`
- Modify: `spec/typescript/exceljs.spec.ts`

- [ ] **Step 1: Write tests for ESM import (`excel.mjs`) and TypeScript definitions for all 6 new features in `spec/typescript/exceljs.spec.ts`**
- [ ] **Step 2: Create `excel.mjs`, update `package.json` (`"exports"` and `"files"`), and update `index.d.ts`**
- [ ] **Step 3: Run `npm run build`, `npm run test:typescript`, ESM import check, `npm run test:unit`, `npm run test:integration`, and `eslint`**
