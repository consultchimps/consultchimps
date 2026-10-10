---
"@consultchimps/xlsx": minor
"@consultchimps/pptx": minor
"@consultchimps/messages": minor
"consultchimps": minor
---

PowerPoint population reads a `.csv` data file (ADR 0007). The file is one
worksheet read by the shared header rule, with the same `csv` options (`--csv-*`
flags) as the sheets commands. A number read by `--csv-numbers` shows as General
would, and a date read by `--csv-dates` as `yyyy-mm-dd`, as in a workbook
ConsultChimps writes. A worksheet name is refused for a CSV file with
`XLSX_CSV_INVALID_OPTION`. `readWorksheetRecords` and
`readWorksheetRecordsBytes` take CSV, and population reports a `csvInputFiles`
metric.
