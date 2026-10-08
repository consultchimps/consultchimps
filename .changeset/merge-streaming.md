---
"@consultchimps/xlsx": minor
---

Merging workbooks now reads each input in pieces and writes the merged workbook
as it is produced (ADR 0006). Merging two workbooks of 150,000 rows peaks at 320
MB instead of 938 MB. The merged workbook is byte for byte what it was.

Add `mergeWorkbookSources`, which merges workbooks read from random-access
sources, such as `blobSource`, into an output written as it is produced.

A damaged input the readers refuse is now refused by merging too, with
`XLSX_READ_FAILED`: a failed CRC check, or a package without its workbook part,
which gave an error with no code before. Merging and splitting refuse a part
compressed with a method they cannot read when the workbook is opened. A package
JSZip read as empty is now read.
