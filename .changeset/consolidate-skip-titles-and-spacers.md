---
"@consultchimps/xlsx": minor
"@consultchimps/messages": patch
---

Skip title rows and spacer columns when reading worksheets. The header row is
now the first row holding more than half as many values as the fullest of the
ten populated rows below it, or at most one value fewer, so a report title, a
merged banner, or a "Prepared by" line above a table four or more columns wide
no longer becomes the first column name. Rows are skipped only when the header
they leave holds nothing but text, or when every skipped row holds a single
value and the header at least three; otherwise, and wherever a title line holds
as many values as the header, the first row holding a value stays the header, as
it always was, so no data row is ever lost to the detection. The inspection's
`headerColumns` metric now counts the kept columns. A column that holds nothing
in the header row or in any row under it is a spacer and never reaches the
output; a column with values under a blank header is kept and named `column_N`
by its position among the columns that were kept. One rule decides this for the
consolidation, the worksheet readers, the worksheet-records reader, and the
inspection, so `sheets inspect` reports the header row and columns a
consolidation will use. A declared `headerRow` is never second-guessed.
Consolidation results carry two new metrics, `skippedTitleRows` and
`skippedSpacerColumns`, and every worksheet report carries the same two counts;
the plain-language result says what was left out.
