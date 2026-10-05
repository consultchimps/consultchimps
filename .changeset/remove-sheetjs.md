---
"@consultchimps/xlsx": patch
---

The published package no longer bundles SheetJS: every read goes through the
streaming reader, and table ranges are read with the package's own helpers.
