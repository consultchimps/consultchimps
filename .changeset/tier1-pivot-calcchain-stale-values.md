---
"@consultchimps/xlsx": minor
---

Close the first three Tier-1 gaps in the workbook split.

- Pivot caches no longer leak other groups' rows into split outputs. A pivot
  cache is a private copy of the source rows that travels inside the package, so
  every pivot table and cache part is now removed from each output, together
  with its relationships, content-type overrides and the workbook's
  `pivotCaches` registry. The result reports how many were removed and warns
  that the pivot has to be rebuilt.
- Values-mode splits no longer bake aggregates computed over removed rows.
  Before the values conversion runs, the cached result of any formula whose A1
  references reach into rows the group does not receive -- including cross-sheet
  references -- is cleared, so the output shows a reported blank cell rather
  than another group's total presented as this one's.
- The calculation chain is kept consistent. Entries naming cells a split deleted
  are dropped, entries whose rows moved are renumbered, and an emptied chain is
  removed along with its relationship and content-type override.
