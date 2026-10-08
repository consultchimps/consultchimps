---
"@consultchimps/xlsx": minor
---

Add a CSV reader for the sheets operations, built on Papa Parse 5.7.0. It reads
a CSV file as a workbook of one worksheet named after the file, in pieces, with
the encoding settled from a byte order mark, valid UTF-8 or a Windows-1252
fallback, the delimiter guessed or chosen, and numbers and dates read only when
asked. No operation takes CSV yet; see ADR 0007.
