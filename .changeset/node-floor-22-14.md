---
"consultchimps": minor
"@consultchimps/core": minor
"@consultchimps/db": minor
"@consultchimps/files": minor
"@consultchimps/messages": minor
"@consultchimps/pdf": minor
"@consultchimps/pptx": minor
"@consultchimps/tabular": minor
"@consultchimps/theme": minor
"@consultchimps/xlsx": minor
---

Raise the supported Node.js floor from 22.0.0 to 22.14.0. Continuous integration
now tests that exact version: pnpm 11 requires 22.13, and better-sqlite3 13
crashes on the 22.13 patch line, so 22.14.0 is the lowest release on which every
published package works.
