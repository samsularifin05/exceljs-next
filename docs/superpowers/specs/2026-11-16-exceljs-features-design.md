# ExcelJS 6-Feature Enhancement Suite Design

## Overview
Design for adding 6 complete features to ExcelJS while preserving 100% backward compatibility with all existing APIs and tests:
1. **Pivot Table Enhancements:** Multiple pivot tables per workbook, multiple aggregation metrics (`sum`, `count`, `average`, `max`, `min`, `product`, `countNums`, `stdDev`, `stdDevP`, `var`, `varP`), multiple `values` fields, and `sourceTable` support.
2. **Streaming Writer & Reader Parity:** `WorksheetWriter.addTable()`, `WorksheetWriter.addImage()`, `WorksheetWriter` extended conditional formatting (`<extLst>`), and `WorksheetReader` `<mergeCell>` parsing (`worksheetReader.mergeCells`).
3. **Column Page Breaks (`colBreaks`):** `column.addPageBreak(top, bottom)`, `worksheet.colBreaks`, `ColBreaksXform` (`<colBreaks>`), and streaming writer support.
4. **Auto-Fit Column Width:** `worksheet.autoFitColumns(options)` calculating optimal column widths from cell contents, headers, and font size.
5. **Workbook Protection:** `workbook.protect(password, options)` and `workbook.unprotect()` with `<workbookProtection>` XML serialization/parsing in `WorkbookXform`.
6. **Native ESM Entry Point:** `excel.mjs` with default and named exports plus `package.json` conditional `"exports"` map (preserving `"main"`, `"browser"`, `"types"`, and deep imports).
