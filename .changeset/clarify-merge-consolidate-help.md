---
"consultchimps": patch
---

Clarify the difference between `sheets merge` and `sheets consolidate` in CLI
help. Each command's description now states the shape of the result up front
(merge keeps every worksheet as its own tab; consolidate stacks all rows into
one combined sheet, matching columns by header), and each command's help now
points to the other command for when that result is what you want.
