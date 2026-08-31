---
"@consultchimps/xlsx": patch
"consultchimps": patch
---

Name a preserved split's outputs after the workbook they actually are. Splitting
a macro-enabled workbook while naming an Excel Table kept the whole source
package — macro project included — but named every output `.xlsx` and reported
the ordinary workbook media type, so the file's contents and its name disagreed
and Excel opened it with a corruption warning. Those outputs are now named
`.xlsm` and carry the macro-enabled media type, exactly as the all-worksheet
split has always done, and a package whose declared type contradicts the name it
arrived under is refused with `XLSX_SPLIT_PACKAGE_TYPE_MISMATCH` before anything
is written — on the preview as well as the run.

A split that rebuilds instead of preserving (`preserveWorkbook: false`, a named
worksheet, or a named range) is unchanged: it writes a fresh ordinary package
from the rows it kept, carries no macro project, and is still `.xlsx` whatever
the source was called.

The CLI carries the same correction:
`consultchimps sheets split <workbook.xlsm> --table <name>` now writes `.xlsm`
files and reports them with the macro-enabled media type, and refuses a workbook
whose package contradicts its name.
