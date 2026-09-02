---
"consultchimps": minor
---

The `sheets unprotect` command now reads its own plain-language result summary
and refuses an output name whose extension contradicts the workbook's type (an
ordinary workbook named `.xlsm`, or a macro-enabled workbook named `.xlsx`)
before writing anything, reporting the stable
`XLSX_UNPROTECT_PACKAGE_TYPE_MISMATCH` reference.
