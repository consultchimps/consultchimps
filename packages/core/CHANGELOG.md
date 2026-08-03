# @consultchimps/core

## 0.4.0

### Minor Changes

- 6564e24: Add byte-level PDF operations for environments without a filesystem,
  such as browsers. `@consultchimps/pdf/bytes` exports `splitPdfBytes`,
  `planSplitPdfBytes`, and `mergePdfsBytes`, which take named in-memory bytes
  and return the produced bytes alongside the same structured result the
  path-based operations report. The entry point's import graph contains no Node
  built-ins, outputs remain byte-deterministic, and cancellation and progress
  work identically. `@consultchimps/core` gains the shared `ByteArtifact` and
  `ByteOperationOutcome` contracts. Existing path-based APIs are unchanged.

## 0.3.0

### Minor Changes

- 5c08c75: Publish typed error-code registries (`FILES_ERRORS`, `PDF_ERRORS`,
  `XLSX_ERRORS`, `PPTX_ERRORS`, with matching `*ErrorCode` unions) so consumers
  can match expected failures without string literals, and make
  `OperationResult` and `OperationPlan` generic over each operation's metric
  names so metric renames become compile-time errors. All runtime values and
  error codes are unchanged; the generics default to `string`, so existing
  consumers keep compiling.

## 0.2.0

### Minor Changes

- c78b35e: Harden the operation APIs for interface consumers. Breaking for
  library users on 0.x: `consolidateWorkbooks`, `splitWorkbookByColumn`,
  `splitPdf`, and `mergePdfs` now take a single options object
  (`inputs`/`input`, `output`/ `outputDirectory`, plus the existing options)
  instead of positional arguments. Every operation now accepts an optional
  `AbortSignal` and a deterministic `onProgress` reporter, and gains a `plan`
  variant (`planConsolidateWorkbooks`, `planSplitWorkbookByColumn`,
  `planSplitPdf`, `planMergePdfs`, `planPopulatePowerPointTemplate`) that
  validates inputs and reports every intended output and collision without
  writing anything. Cancellation raises a stable `OPERATION_ABORTED` error and
  never modifies source files. CLI behavior is unchanged.
