---
"@consultchimps/xlsx": minor
"@consultchimps/messages": minor
"consultchimps": minor
---

Write CSV (ADR 0007). Consolidate writes a CSV file when the output is named
`.csv` or with `--output-format csv`, and split writes CSV files for a CSV input
or with `--output-format csv`, splitting compactly. A CSV output uses RFC 4180
quoting, commas and CRLF, is UTF-8 with a byte order mark unless
`--csv-bom false`, writes dates as ISO text, and prefixes an apostrophe to text
that would start a formula, leaving signed numbers alone. A merge named `.csv`,
or a format that contradicts the output's name, is refused with
`XLSX_OUTPUT_FORMAT_INVALID`.
