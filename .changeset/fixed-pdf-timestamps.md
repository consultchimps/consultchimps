---
"@consultchimps/pdf": patch
---

Generated PDFs now carry a fixed creation and modification timestamp instead of
the current time, so identical inputs and options produce byte-identical split
and merge outputs, matching the determinism promise the other document formats
already keep.
