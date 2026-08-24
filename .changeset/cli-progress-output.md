---
"consultchimps": minor
"@consultchimps/xlsx": minor
---

Show live progress on stderr during long-running CLI commands.
`sheets consolidate`, `sheets merge`, `sheets split`, `pdf split`, `pdf merge`,
and `pptx populate` now report the current stage and item count while they work
(for example `Reading workbooks 3/14: report.xlsx`), so large jobs no longer sit
silent until the final report. In an interactive terminal the line updates in
place; when output is redirected, plain lines are printed instead. Progress goes
only to stderr, never interleaves with the final report, and is fully suppressed
under `--json`.

`mergeWorkbooks` in `@consultchimps/xlsx` now accepts the standard operation
controls (`onProgress`, `signal`), emitting `merging-inputs` events per input
workbook and a final `writing-output` event, matching `consolidateWorkbooks` and
the in-memory `mergeWorkbooksBytes`.
