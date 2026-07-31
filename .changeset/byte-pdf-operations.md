---
"@consultchimps/core": minor
"@consultchimps/pdf": minor
---

Add byte-level PDF operations for environments without a filesystem, such as
browsers. `@consultchimps/pdf/bytes` exports `splitPdfBytes`,
`planSplitPdfBytes`, and `mergePdfsBytes`, which take named in-memory bytes
and return the produced bytes alongside the same structured result the
path-based operations report. The entry point's import graph contains no Node
built-ins, outputs remain byte-deterministic, and cancellation and progress
work identically. `@consultchimps/core` gains the shared `ByteArtifact` and
`ByteOperationOutcome` contracts. Existing path-based APIs are unchanged.
