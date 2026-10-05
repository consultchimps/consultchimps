---
"@consultchimps/xlsx": minor
"consultchimps": minor
---

A date cell read from a workbook is written back as an Excel date, not as ISO
text. Consolidate, a compact split and `writeTable` write it as a 1900-system
serial formatted `yyyy-mm-dd`, or `yyyy-mm-dd hh:mm:ss` when it carries a time,
so the output sorts, filters and calculates on it. Dates from 1904-system inputs
are recounted, and a date before 1900 stays text. A `writeTable` caller that
passes a string in that exact timestamp spelling now gets a date cell too.
