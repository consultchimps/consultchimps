---
"consultchimps": patch
---

`sheets inspect`, `sheets split` and `sheets merge` now read each workbook in
pieces and write outputs as they go (ADR 0006), so large workbooks use far less
memory. Splitting a workbook of 150,000 rows no longer fails with "Maximum call
stack size exceeded". A failed CRC check or malformed XML in a part a command
reads is now refused, as the other readers already did, instead of being
described, split or merged. `sheets inspect` decodes `_x000D_` and similar
escapes, leaves phonetic text out of headers, and lists sample values in row
order, as the other commands already did.
