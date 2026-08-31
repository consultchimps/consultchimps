---
"@consultchimps/xlsx": patch
---

Export the conformance contract: `CONTRACT`, `TRACKED_STRUCTURES`, `OPERATIONS`,
the `UNDECIDED_*` records, and their types now ship from the package root. The
table states, per workbook structure and operation, whether the package
preserves it, rewrites it so it stays valid, removes it with a warning, or
refuses — the same table the corpus tests enforce — so a caller can tell people
what an operation will do before running it, and the documentation site can
generate its preservation matrix instead of restating it in prose.
