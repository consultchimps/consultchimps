---
"@consultchimps/xlsx": minor
"@consultchimps/pptx": minor
"consultchimps": minor
---

A formula cell with no cached value, from a workbook that was saved without
being calculated, is no longer read as blank in silence. Consolidate, split,
`merge --values`, `sheets inspect` and PowerPoint populate count such cells in a
new `formulaCellsWithoutCachedValues` metric and name them in one warning, up to
ten `Sheet!B4` locations, so the user knows to open and save the file in Excel.
`readWorksheetRecords` returns them as `uncachedFormulas`. Nothing computes the
missing value, and a preserving split or merge still carries the formula.
