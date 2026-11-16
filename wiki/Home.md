# Welcome to the `exceljs-next` Wiki

[`exceljs-next`](https://github.com/samsularifin05/exceljs-next) is a maintained, modernized, and security-hardened fork of `exceljs@4.4.0` supporting **Node.js 16–22+**, **Native ESM**, **TypeScript**, and **Browsers** with **0 npm audit vulnerabilities**.

---

## Tutorial & Documentation Pages

### Getting Started & New `exceljs-next` Features
1. **[Installation & Quick Start](Installation-and-Quick-Start)** — Install via `npm` / `pnpm`, drop-in `exceljs` alias, CommonJS, Native ESM, and TypeScript configuration.
2. **[File Encryption & Protection](File-Encryption-and-Protection)** — Complete guide to File-Open Password (`MS-OFFCRYPTO` AES-256), Worksheet Cell Protection (`ws.protect()`), and Workbook Structure Protection (`workbook.protect()`).
3. **[Native Excel Charts & Data Labels](Charts-and-Data-Labels)** — Create Column, Bar, Line, Area, Pie, and Doughnut charts with custom colors, `valAxis` number formats, and percentage data labels.
4. **[Template Engine (`renderTemplate`)](Template-Engine)** — Populate `.xlsx` templates with `{{placeholder}}` variables and automatic array loop row expansion.
5. **[JSON Import/Export & Auto-Fit](JSON-and-AutoFit)** — Use `ws.fromJSON()`, `ws.toJSON()`, `ws.autoFitColumns()`, and `ws.autoFitRows()`.
6. **[Pivot Tables & Streaming](Pivot-Tables-and-Streaming)** — Multi-Pivot Tables, 11 aggregation metrics, `sourceTable`, and Streaming `WorkbookWriter` / `WorkbookReader`.

### Core Spreadsheet Guides
7. **[Workbook & Worksheet Basics](Workbook-and-Worksheet-Basics)** — Workbook metadata, `views` (freeze/split), columns, rows, cells, rich text, hyperlinks, and merge cells.
8. **[Styling & Number Formats](Styling-and-Number-Formats)** — Fonts, pattern & gradient fills, alignment, borders, and `numFmt`.
9. **[Formulas & Data Validation](Formulas-and-Data-Validation)** — Standard formulas, shared formulas (`fillFormula`), and dropdown/number data validation.
10. **[Tables, Conditional Formatting, & Media](Tables-Conditional-Formatting-and-Media)** — Excel Tables (`addTable`), Conditional Formatting, Images, Comments, and Page Breaks.
11. **[CSV & Browser Usage](CSV-and-Browser-Usage)** — Reading/writing CSV files and using `dist/exceljs.min.js` in the browser.
12. **[Page Setup, Auto Filters, Outlines, & Defined Names](Page-Setup-AutoFilter-and-Advanced)** — Print setup, margins, headers/footers, autoFilter, outline grouping, defined names, and `ValueType`.

---

## Interactive Browser Playground

Want to test features visually without writing code first? Open [`demo.html`](https://github.com/samsularifin05/exceljs-next/blob/master/demo.html) locally in your browser to generate `.xlsx` files, inspect sheets, and copy live code snippets.
