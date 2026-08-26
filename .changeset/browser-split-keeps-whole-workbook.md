---
"@consultchimps/xlsx": minor
---

Split whole workbooks by default without a filesystem, so the in-browser
splitter now works the way the command line does.

`splitWorkbookBytes` and `planSplitWorkbookBytes` previously read a single
worksheet and wrote compact, data-only workbooks. They now collect the values of
the chosen column across every worksheet and give back one complete copy of the
source workbook per value:

- a worksheet that carries the column keeps its header and only that value's
  rows;
- a worksheet that does not carry the column is copied through untouched; and
- sheet order and visibility, formatting, merged cells, conditional formatting,
  data validation, hyperlinks, comments, images, charts, defined names, and the
  macro parts of an `.xlsm` file all survive.

`.xlsm` workbooks are accepted as input and produce `.xlsm` outputs.

Strict matching is available, so case, surrounding whitespace, and value type
can be kept distinct instead of `North`, `north`, and `North ` becoming one
workbook.

Results report more of what happened: how many worksheets were filtered, how
many were copied unchanged, and, for each output workbook, the rows kept and
removed per worksheet. The values-only, pivot-cache and stale-cached-total
warnings match the ones the command line reports.

The single-source split is still available for narrower jobs: name an Excel
Table, a named range, or a worksheet, or turn off whole-workbook preservation
for compact data-only outputs.
