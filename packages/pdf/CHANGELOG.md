# @consultchimps/pdf

## 0.4.3

### Patch Changes

- Updated dependencies [846a8bd]
  - @consultchimps/files@0.4.0

## 0.4.2

### Patch Changes

- 32973f7: Declare support for Node.js 22 and later (`engines.node: ">=22.0.0"`
  instead of `24.x`), so the toolkit installs and runs in environments that ship
  the previous LTS line. CI now validates the runtime on Node 22.16, the latest
  22, and 26 alongside the full Node 24 verification. The CLI additionally ships
  a standalone `consultchimps.mjs` bundle on each GitHub release: one file that
  runs with `node consultchimps.mjs` and needs no npm access at all.
- Updated dependencies [32973f7]
  - @consultchimps/core@0.5.1
  - @consultchimps/files@0.3.3

## 0.4.1

### Patch Changes

- 1f759eb: Centralize portable filename sanitization in `@consultchimps/core`,
  which now exports `safeNameFragment` and `truncateToUtf8Bytes`. The PDF and
  Excel operations share that single implementation instead of keeping their own
  copies, so every generated output name follows the same rules.

  Excel splitting previously capped a group value at 80 code points rather than
  80 UTF-8 bytes, so a long non-ASCII group value could produce an output
  filename far past common 255-byte filename limits. It is now capped by encoded
  size, truncating only at code point boundaries. Split filenames are unchanged
  for ASCII group values; a non-ASCII group value longer than 80 UTF-8 bytes is
  now shortened.

- Updated dependencies [1f759eb]
  - @consultchimps/core@0.5.0
  - @consultchimps/files@0.3.2

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

### Patch Changes

- Updated dependencies [6564e24]
  - @consultchimps/core@0.4.0
  - @consultchimps/files@0.3.1

## 0.3.0

### Minor Changes

- 5c08c75: Publish typed error-code registries (`FILES_ERRORS`, `PDF_ERRORS`,
  `XLSX_ERRORS`, `PPTX_ERRORS`, with matching `*ErrorCode` unions) so consumers
  can match expected failures without string literals, and make
  `OperationResult` and `OperationPlan` generic over each operation's metric
  names so metric renames become compile-time errors. All runtime values and
  error codes are unchanged; the generics default to `string`, so existing
  consumers keep compiling.

### Patch Changes

- 849bf37: Generated PDFs now carry a fixed creation and modification timestamp
  instead of the current time, so identical inputs and options produce
  byte-identical split and merge outputs, matching the determinism promise the
  other document formats already keep.
- Updated dependencies [5c08c75]
  - @consultchimps/core@0.3.0
  - @consultchimps/files@0.3.0

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

### Patch Changes

- Updated dependencies [c78b35e]
  - @consultchimps/core@0.2.0
  - @consultchimps/files@0.2.0

## 0.1.1

### Patch Changes

- Updated dependencies [6870be2]
  - @consultchimps/files@0.1.1
