---
"@consultchimps/xlsx": minor
"@consultchimps/messages": patch
---

Skip title rows and spacer columns when reading worksheets. The header row is
now the first row holding more than half as many values as the fullest of the
ten populated rows below it, or at most one value fewer, so a report title, a
merged banner, or a "Prepared by" line above a table four or more columns wide
no longer becomes the first column name. A row is skipped only on evidence that
it cannot be the header of the block below it, a blank row between the two or
fewer than a third as many values as the header holds, and never when the
skipped rows form a table of their own; otherwise the first row holding a value
stays the header, as it always was, so no data row is ever lost to the
detection. The inspection's `headerColumns` metric now counts the kept columns.
A column that holds nothing in the header row or in any row under it is a spacer
and never reaches the output; a column with values under a blank header is kept
and named `column_N` by its position among the columns that were kept. One rule
decides this for the consolidation, the worksheet readers, the worksheet-records
reader, and the inspection, so `sheets inspect` reports the header row and
columns a consolidation will use, spelled the same way whatever the header
cell's type; an error cell in a header row now names its column `#REF!` rather
than Excel's internal code for it. A split looks for its column on the detected
header row before searching the whole worksheet, so a title line that repeats
the column's name no longer captures the split. A declared `headerRow` is never
second-guessed. Consolidation results carry two new metrics, `skippedTitleRows`
and `skippedSpacerColumns`, and every worksheet report carries the same two
counts; the plain-language result says what was left out.
