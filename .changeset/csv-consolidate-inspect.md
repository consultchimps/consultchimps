---
"@consultchimps/xlsx": minor
"@consultchimps/messages": minor
"consultchimps": minor
---

Consolidate and inspect take CSV files beside workbooks (ADR 0007). Each CSV
file is one worksheet named after the file, read in pieces with its encoding and
delimiter detected, and every field stays text unless `csv` options
(`--csv-numbers`, `--csv-decimal`, `--csv-thousands`, `--csv-dates`) ask for
numbers or dates; `--csv-delimiter` overrides the guess and `--csv-encoding` the
encoding, except that a byte order mark wins over a contradicting choice, with a
warning. Consolidation reports a `csvInputFiles` metric and refuses a table
larger than a worksheet holds with `XLSX_OUTPUT_TOO_LARGE`. An inspection of a
CSV file reports the encoding and delimiter it used.
