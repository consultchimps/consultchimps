---
"@consultchimps/core": minor
"@consultchimps/files": minor
"@consultchimps/pdf": minor
"@consultchimps/pptx": minor
"@consultchimps/xlsx": minor
"@consultchimps/messages": patch
"consultchimps": patch
---

Publish typed error-code registries (`FILES_ERRORS`, `PDF_ERRORS`,
`XLSX_ERRORS`, `PPTX_ERRORS`, with matching `*ErrorCode` unions) so consumers
can match expected failures without string literals, and make `OperationResult`
and `OperationPlan` generic over each operation's metric names so metric renames
become compile-time errors. All runtime values and error codes are unchanged;
the generics default to `string`, so existing consumers keep compiling.
