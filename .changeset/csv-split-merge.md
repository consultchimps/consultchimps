---
"@consultchimps/xlsx": minor
"@consultchimps/messages": minor
"consultchimps": minor
---

Split and merge take CSV files (ADR 0007). A CSV file splits into compact
workbooks, since it has no workbook to keep; asking to keep the workbook is
refused with `XLSX_SPLIT_CSV_PRESERVE`. A merge makes each CSV file one tab
named after the file, its rows copied as they are, with dates read by
`--csv-dates` stored as Excel dates. Merges report a `csvInputFiles` metric. The
`--csv-*` options apply to `sheets split` and `sheets merge`.
