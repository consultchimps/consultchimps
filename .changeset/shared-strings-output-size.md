---
"@consultchimps/xlsx": patch
---

Write consolidated and rebuilt split workbooks with a shared-strings table
instead of per-cell strings, matching how Excel stores text. Outputs with
repetitive text now serialize noticeably smaller — roughly 13-20% on synthetic
benchmarks, with the biggest gains when repeated values are spread far apart in
large workbooks.
