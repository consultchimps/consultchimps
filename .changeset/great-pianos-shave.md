---
"@consultchimps/messages": minor
"@consultchimps/xlsx": minor
"consultchimps": minor
---

Consolidation can now fold columns that are named differently into one column
each, using the versioned JSON column mapping document.

`sheets consolidate` gains `--map <file>`, which applies a mapping before the
rows are stacked, and `--suggest-map <file>`, which writes a draft mapping built
from the headers that were read and still writes the consolidated workbook. A
draft is never applied for you, it goes through the same never-overwrite rule as
any other output, and the two options cannot be combined in one run.

A column no mapping entry claims keeps its own name and is reported as a
warning. Two columns of one worksheet folding into one canonical column stop the
run before anything is written. A declared date coercion reads text: a column
holding a number, or a value Excel already stores as a date, is refused by name
rather than read as a date serial, because which day a serial counts from
belongs to the workbook rather than to the cell.

`consolidateWorkbooks` takes `mappingFile` and `suggestMappingOutput`, and
`consolidateWorkbooksBytes` takes a parsed `mapping` and `suggestMapping`. Both
return the drafted mapping on the result, and both report two new metrics,
`unmappedColumns` and `suggestedColumns`. Result explanations name the columns a
mapping did not claim, point at a drafted mapping among the created files, and
give mapping failures their own recovery steps.
