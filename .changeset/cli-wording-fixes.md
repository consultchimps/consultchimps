---
"@consultchimps/files": minor
"@consultchimps/xlsx": minor
"@consultchimps/pptx": minor
"@consultchimps/messages": minor
"consultchimps": minor
---

Merge keeps the order you give. `sheets merge` and `pdf merge` use named files
in the order given, and a folder or pattern adds its matches alphabetically at
its place; `discoverFiles` takes `order: "given"` for this, and sorts paths with
a fixed English collation so Linux and Windows agree.

Output names what was read. CSV inputs and outputs are no longer called Excel
workbooks in summaries, progress lines or help; `describeWorkbook` reports a
`csvInputFiles` metric; every metric has a plain-language label; the next steps
point at the files listed above them; `pptx populate` progress names slides as
"slide 1" rather than part names; `pptx inspect-template` drops the always-empty
split-run line, since split-run placeholders are supported.

Usage mistakes report `CLI_USAGE` rather than no code: a number option given
anything but a whole number from 1 (`--header-row 1.5` used to read row 1, and
the db `--limit` options accepted any text), `-o` with `--output-dir`, and a
pattern matching several files for a one-file command (which `pdf split` and
`sheets unprotect` used to resolve silently to the first match). A file named in
one letter case and matched by a pattern in another is used once. `db` help
descriptions are lowercase like the others.
