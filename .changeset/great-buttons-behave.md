---
"@consultchimps/xlsx": minor
---

Consolidate workbooks without a filesystem. `@consultchimps/xlsx/bytes` now
exports `consolidateWorkbooksBytes`, which takes named in-memory workbooks and
returns one combined workbook's bytes alongside the same structured result the
path-based `consolidateWorkbooks` reports. It accepts the same
`normalizeHeaders`, `addSourceColumns`, worksheet-selection, and header-row
options, reports progress, and can be cancelled, and it produces byte-identical
output to the path-based operation for the same workbooks and options — so a
browser and the command line agree exactly.
