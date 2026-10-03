---
"@consultchimps/xlsx": minor
---

Write consolidated and split workbooks with a streaming worksheet writer built
on fflate (ADR 0006), instead of building each one in memory through SheetJS.
The writer deflates the worksheet XML row by row, so it no longer holds a copy
of every row as an array, then as one cell object per value, before the file is
written.

The workbook's bytes change once, deliberately. Text is stored in each cell
rather than in a shared-strings table, and the package drops the theme and
metadata parts it never used. Every cell value, column width, the header row,
the filter, and the used range are unchanged, the CLI and the browser still
produce byte-identical files, and identical inputs still give identical bytes in
every time zone.

Text that looks like Excel's own escape for a control character, such as
`_x0041_`, now reads back exactly; SheetJS wrote it unescaped, so every reader
turned it into a different character. A worksheet name Excel cannot open now
fails with `XLSX_INVALID_SHEET_NAME` instead of an unexplained error.
