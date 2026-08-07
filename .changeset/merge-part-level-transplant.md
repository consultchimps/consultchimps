---
"@consultchimps/xlsx": minor
---

The workbook merge now preserves Excel Tables, defined names, conditional
formatting, data validation, cell comments, styles and number formats at the
package level instead of rebuilding each worksheet through a spreadsheet
library.

`mergeWorkbooks` and `mergeWorkbooksBytes` keep their signatures, their metrics
and their error codes. What changed is the engine underneath: the first input
now seeds the output package and every later input's worksheet parts are copied
into it with the parts they depend on, so a structure survives unless carrying
it would be wrong.

- **Preserved**: merged cells, conditional formatting, data validation,
  hyperlinks, comments and their drawings, Excel Tables including totals rows,
  cell styles and number formats, and shared, array and uncached formulas.
- **Repaired**: shared-string and style indexes are remapped into one merged
  table per workbook (identical entries collapse); duplicate Excel Table names
  and workbook-scoped defined names take a numeric suffix, with a warning
  listing every rename; formulas that named a renamed worksheet or table -
  structured references included - are rewritten to follow it.
- **Removed, each with a warning**: pivot tables and their caches, external
  links, and a macro project that cannot travel. A macro project is carried only
  when a single input has one and the output is named `.xlsm`; the byte surface
  now keeps an `.xlsm` output name a caller asks for, and reports the
  macro-enabled media type for it.
- The calculation chain is dropped without a warning and the merged workbook
  asks Excel to recalculate on open, because a chain is a derived index keyed by
  sheet ids that every transplanted worksheet changes.
