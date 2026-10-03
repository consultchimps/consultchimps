---
"@consultchimps/xlsx": minor
"@consultchimps/tabular": minor
---

Consolidate workbooks in two streaming passes (ADR 0006). The first pass reads
each worksheet's header rows, spacer columns, column union, mapping, and column
widths without keeping its rows; the second reads the rows again and writes them
straight to the output. Memory now follows the largest worksheet rather than the
sum of the inputs: ten workbooks of 15,000 rows each peak at a fraction of the
memory they needed before. The command line writes the output to a staging file
beside the destination and moves it into place only when every row is written,
so a failed or cancelled run leaves nothing behind.

Every cell value is read as before, with three deliberate changes. An error cell
is written as the error it holds, `#DIV/0!`, instead of Excel's internal number
for it. A carriage return stored before a line feed is kept rather than dropped.
A formula's cached text is unescaped once, as all other text is, so text that
looks like an escape reads back as written. Because an error cell no longer
reads as a number, a mapping that coerces its column to a number now refuses it
rather than writing Excel's internal code as an amount. Consolidation no longer
parses each input twice, once for values and once for dates.

`@consultchimps/tabular` exports `planTableUnion`, the column plan behind
`unionTables`, for callers that stack rows as they stream.
