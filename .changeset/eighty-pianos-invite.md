---
"consultchimps": minor
---

Add `consultchimps sheets inspect <workbook>`, which describes an Excel workbook
without creating or changing anything. The report lists each described worksheet
with its visibility, used range, resolved header row, and data row count, every
header with a few of the values stored beneath it, and the Excel Tables and
named ranges those worksheets contain, followed by the usual plain-language
explanation of the result. `--sheet`, `--header-row`, and `--hidden` select what
is described, and `--samples` sets how many sample values each column reports,
from 0 to 5. Under `--json` the envelope carries the whole inspection outcome:
the description beside the operation result.
