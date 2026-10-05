---
"@consultchimps/xlsx": minor
"consultchimps": patch
---

`consolidateWorkbookSources` in `@consultchimps/xlsx/bytes` consolidates
workbooks read in pieces into an output written chunk by chunk, so a browser can
consolidate large files without holding them whole. `blobSource` reads a browser
`File` or `Blob` through `Blob.slice`. The output bytes match
`consolidateWorkbooksBytes` and the command line.

A consolidation whose output cannot be written now fails with that write error,
rather than reporting that an input could not be read.
