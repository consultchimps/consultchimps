---
"consultchimps": minor
"@consultchimps/messages": minor
---

Add `consultchimps pbi export <input>`, which writes `workbook.xlsx` and
`manifest.json` from a Power BI file without opening Power BI Desktop. It takes
`-o` for the output folder, `--include-hidden` for the tables a model marks
hidden, `--max-memory` for the ceiling the export may reserve, and `--force` to
replace existing outputs, and it lists the tables exported, what was left out
and why, and the readings that have no verified example.
`@consultchimps/messages` renders every Power BI code in plain language for both
the generic and the command-line vocabularies.
