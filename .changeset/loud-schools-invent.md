---
"@consultchimps/xlsx": minor
---

Add workbook inspection: `describeWorkbook` and `describeWorkbookBytes` report a
workbook's structure without creating anything — worksheets with their
visibility and dimensions, the header row an operation would actually use, Excel
Tables, named ranges, and up to five distinct sample values per column. The
outcome pairs that description with a structured `sheets.inspect` result
carrying counts as metrics, no artifacts, and warnings for hidden worksheets
left out and worksheets with no header row.

The bytes surface also gains `readWorkbookExcelTablesBytes` and
`readWorkbookNamedRangesBytes`, the byte twins of the path-based Excel Table and
named-range readers, so its `WorkbookExcelTable` and `WorkbookNamedRange` type
exports now describe values that surface can actually produce.
