---
"@consultchimps/xlsx": minor
"@consultchimps/pptx": patch
"consultchimps": patch
---

Splitting a workbook now reads it in pieces and writes each output as it is
produced, in every mode (ADR 0006), and no longer fails with "Maximum call stack
size exceeded" on large worksheets. Splitting 150,000 rows by a column, keeping
the workbook, finishes at 335 MB instead of failing after 1.9 GB. Outputs are
byte for byte what they were.

Add `splitWorkbookSource` and `planSplitWorkbookSource`, which split a workbook
read from a random-access source such as `blobSource` into outputs written as
they are produced.

A damaged workbook the readers refuse is now refused by splitting too: a failed
CRC check, or a malformed row or cell reference, which gave an error with no
code before. A package JSZip read as empty is now read.
