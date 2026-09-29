---
"consultchimps": minor
"@consultchimps/core": minor
"@consultchimps/xlsx": minor
---

Keep a local record of every CLI run. Each command that does work writes one
JSON Lines file with the machine it ran on, every progress event with its time,
memory sampled every second, and how it ended, so a slow or failed run can be
explained afterwards. Records hold counts, sizes, and timings only;
`--log-names` adds file names, progress details, and error messages. `--no-log`
or `CONSULTCHIMPS_LOG=off` turns recording off, `--profile` also saves a CPU
profile, and `consultchimps logs`, `logs show`, and `logs path` read the
records.

`@consultchimps/core` gains the record types, `summarizeTaskLog`, and an
optional `measures` field on `OperationProgress`. Consolidation in
`@consultchimps/xlsx` reports the tables, rows, and columns each workbook
yielded through it.
