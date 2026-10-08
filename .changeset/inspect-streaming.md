---
"@consultchimps/xlsx": minor
---

`describeWorkbook` and `describeWorkbookBytes` now read through the streaming
reader the table readers and consolidation use (ADR 0006), one worksheet at a
time, keeping only the rows the header rule needs and a bounded sample. The
command line reads the file in pieces. Inspecting 150,000 rows peaks at 185 MB
instead of 2.3 GB.

Descriptions are unchanged for ordinary workbooks. Where inspection read a cell
differently from the readers, it now reads it as they do: `_x000D_` and other
escapes are decoded, phonetic text is left out of a header, and samples follow
row order. A damaged package the readers refuse, such as a failed CRC check or
malformed XML, is now refused by inspection too, instead of described.

An Excel Table on a worksheet whose name holds an escape is now found on that
worksheet, by inspection and by `readWorkbookExcelTables`.
