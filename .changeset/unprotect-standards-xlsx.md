---
"@consultchimps/xlsx": minor
---

Bring the workbook unprotect operation up to the package standard. The XML edit
now runs on the package layer through a new
`WorkbookPackage.removeEmptyElements` seam instead of a regex inside the
operation, unprotect gains a declared column in the conformance contract (every
tracked structure is `preserve`, held up by a new corpus), and unprotect now
refuses an output name whose extension contradicts the workbook's declared type
with the new stable error `XLSX_UNPROTECT_PACKAGE_TYPE_MISMATCH`, before writing
anything.
