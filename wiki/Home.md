# Welcome to the `exceljs-next` Wiki

[`exceljs-next`](https://github.com/samsularifin05/exceljs-next) is a maintained, modernized, and security-hardened fork of `exceljs@4.4.0` supporting **Node.js 16–22+**, **Native ESM**, **TypeScript**, and **Browsers** with **0 npm audit vulnerabilities**.

---

## Tutorial & Documentation Pages

1. **[Installation & Quick Start](Installation-and-Quick-Start)** — Install via `npm` / `pnpm`, drop-in `exceljs` alias, CommonJS, Native ESM, and TypeScript configuration.
2. **[File Encryption & Protection](File-Encryption-and-Protection)** — Complete guide to all 3 protection levels:
   - File-Open Password (`MS-OFFCRYPTO` AES-256)
   - Worksheet Cell Protection (`ws.protect()`)
   - Workbook Structure Protection (`workbook.protect()`)
3. **[Native Excel Charts & Data Labels](Charts-and-Data-Labels)** — Create Column, Bar, Line, Area, Pie, and Doughnut charts with custom colors, `valAxis` number formats, and percentage data labels.
4. **[Template Engine (`renderTemplate`)](Template-Engine)** — Populate `.xlsx` templates with `{{placeholder}}` variables and automatic array loop row expansion.
5. **[JSON Import/Export & Auto-Fit](JSON-and-AutoFit)** — Use `ws.fromJSON()`, `ws.toJSON()`, `ws.autoFitColumns()`, and `ws.autoFitRows()`.
6. **[Pivot Tables & Streaming](Pivot-Tables-and-Streaming)** — Multi-Pivot Tables, 11 aggregation metrics, `sourceTable`, and Streaming `WorkbookWriter` / `WorkbookReader`.

---

## Interactive Browser Playground

Want to test features visually without writing code first? Open [`demo.html`](https://github.com/samsularifin05/exceljs-next/blob/master/demo.html) locally in your browser to generate `.xlsx` files, inspect sheets, and copy live code snippets.
