---
"consultchimps": patch
"@consultchimps/core": patch
"@consultchimps/files": patch
"@consultchimps/messages": patch
"@consultchimps/pdf": patch
"@consultchimps/pptx": patch
"@consultchimps/tabular": patch
"@consultchimps/xlsx": patch
---

Allow installation on Node.js 22 and newer. The packages were pinned to
`node 24.x`, which made strict package managers refuse to install them on the
still-supported Node 22 line even though nothing in the code requires Node 24.
