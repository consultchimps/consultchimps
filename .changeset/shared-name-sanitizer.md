---
"@consultchimps/core": minor
"@consultchimps/pdf": patch
"@consultchimps/xlsx": patch
---

Centralize portable filename sanitization in `@consultchimps/core`, which now
exports `safeNameFragment` and `truncateToUtf8Bytes`. The PDF and Excel
operations share that single implementation instead of keeping their own copies,
so every generated output name follows the same rules.

Excel splitting previously capped a group value at 80 code points rather than 80
UTF-8 bytes, so a long non-ASCII group value could produce an output filename
far past common 255-byte filename limits. It is now capped by encoded size,
truncating only at code point boundaries. Split filenames are unchanged for
ASCII group values; a non-ASCII group value longer than 80 UTF-8 bytes is now
shortened.
