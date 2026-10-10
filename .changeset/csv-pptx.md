---
"@consultchimps/xlsx": minor
"@consultchimps/pptx": minor
"@consultchimps/messages": minor
"consultchimps": minor
---

PowerPoint population reads a `.csv` data file (ADR 0007). The file is one
worksheet read by the shared header rule, with the same `csv` options (`--csv-*`
flags) as the sheets commands. A number read by `--csv-numbers` shows every
digit with a point as its decimal mark, a date read by `--csv-dates` shows as
`yyyy-mm-dd`, and header names stay as written. A worksheet name is refused for
a CSV file with `XLSX_CSV_INVALID_OPTION`. `readWorksheetRecords` and
`readWorksheetRecordsBytes` take CSV, and population reports a `csvInputFiles`
metric.
