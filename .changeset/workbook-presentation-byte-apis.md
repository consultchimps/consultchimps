---
"@consultchimps/pptx": minor
"@consultchimps/xlsx": minor
---

Add byte-level entry points so Excel and PowerPoint operations can run without a
filesystem, such as in a browser.

`@consultchimps/xlsx/bytes` exports `splitWorkbookBytes`,
`planSplitWorkbookBytes`, `mergeWorkbooksBytes`, and
`readWorksheetRecordsBytes`. The split accepts the worksheet, Excel Table, named
range, header row, blank-value, and workbook-preserving options of the
path-based split, including the formatting-preserving Excel Table split; the
merge keeps every worksheet's cells and formatting, resolves tab name
collisions, and reports hidden source worksheets exactly as the path-based merge
does.

`@consultchimps/pptx/bytes` exports `populatePresentationBytes`,
`planPopulatePresentationBytes`, and `inspectPresentationBytes`. The population
reads its records either from an in-memory array or from workbook bytes, and the
new `PPTX_INVALID_DATA_SOURCE` error reports a call that supplies both or
neither.

Byte operations take named in-memory inputs, return the produced bytes with the
same structured result the path-based operations report, sanitize every output
name into a portable filename, support progress reporting and cancellation, and
produce nothing when cancelled.

Generated workbooks and presentations are now byte-identical for identical
inputs: rewritten Open XML parts carry a fixed timestamp instead of the current
time, and packages no longer gain directory entries the source package did not
have. This also fixes clock-dependent bytes in the path-based worksheet merge,
the workbook-preserving split, and PowerPoint population.
