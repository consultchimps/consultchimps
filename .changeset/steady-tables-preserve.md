---
"@consultchimps/xlsx": minor
"consultchimps": minor
---

Excel Table splits now preserve the complete source workbook by default, so
every output opens exactly like the prepared original (zoom, cursor position,
cover sheets, styles, and content outside the table all carry over). Pass
`preserveWorkbook: false` or the new `--no-preserve-workbook` flag for the
previous compact data-only outputs. Behavior change for existing table splits
that relied on the compact default; plain worksheet splits are unchanged.

Preserved splits now refuse to relocate cells whose formulas depend on their
position (A1-style references, shared, or array formulas) with a stable
`XLSX_SPLIT_PRESERVE_FORMULA` error instead of silently producing formulas that
point at the wrong rows; structured table references such as `[@Amount]` remain
fully supported.

Workbook named ranges are now supported as a data source: select one with the
new `range` option or `--range <name>` flag, inspect them with the new
`readWorkbookNamedRanges` export, and prefer sources in the order Excel Table,
named range, then full worksheet range. Named-range splits always produce
compact outputs and reject `headerRow` and `preserveWorkbook`.
