---
"@consultchimps/messages": minor
---

Explain a workbook unprotect result in plain language instead of the generic
fallback. The new `sheets.unprotect` entry reports how many worksheet and
workbook-structure protections were removed, including the case where there was
nothing to remove, and a macro-enabled `.xlsm` output is now named as an Excel
workbook rather than a bare file.
