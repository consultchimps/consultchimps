# @consultchimps/pptx

## 0.6.2

### Patch Changes

- Updated dependencies [e716dc7]
- Updated dependencies [76ec2b3]
  - @consultchimps/xlsx@0.14.0

## 0.6.1

### Patch Changes

- 32973f7: Declare support for Node.js 22 and later (`engines.node: ">=22.0.0"`
  instead of `24.x`), so the toolkit installs and runs in environments that ship
  the previous LTS line. CI now validates the runtime on Node 22.16, the latest
  22, and 26 alongside the full Node 24 verification. The CLI additionally ships
  a standalone `consultchimps.mjs` bundle on each GitHub release: one file that
  runs with `node consultchimps.mjs` and needs no npm access at all.
- Updated dependencies [cef85f7]
- Updated dependencies [32973f7]
  - @consultchimps/xlsx@0.13.0
  - @consultchimps/core@0.5.1
  - @consultchimps/files@0.3.3

## 0.6.0

### Minor Changes

- cc4d06c: Report a PowerPoint template inspection as a structured operation
  result. `inspectPresentationOutcomeBytes()` in `@consultchimps/pptx/bytes`
  returns the placeholder report together with the same `OperationResult` every
  completed ConsultChimps operation returns: the slide's counts as metrics, no
  artifacts, and one plain-language warning for each condition that would make a
  population refuse the template — malformed placeholder braces, placeholders
  outside a supported text shape, placeholders split across text runs, and a
  slide with no usable placeholders at all. Identical templates and options
  produce an identical result. The existing `inspectPresentationBytes()` is
  unchanged and still returns the placeholder report on its own.

  `planPopulatePresentationBytes()` now honours the `signal` its options already
  accepted, and `inspectPresentationBytes()` accepts `signal` and `onProgress`.
  Both read whole packages before they can answer, so a caller that has moved on
  — a page replanning after a keystroke, or inspecting a different slide — can
  stop that work rather than only discard its answer, and an inspection reports
  which of its two package reads a large deck is currently in.

### Patch Changes

- Updated dependencies [6c683ec]
- Updated dependencies [aabc58b]
- Updated dependencies [c7ccb8d]
  - @consultchimps/xlsx@0.12.0

## 0.5.4

### Patch Changes

- Updated dependencies [a969491]
- Updated dependencies [98c77a7]
- Updated dependencies [d1f8524]
  - @consultchimps/xlsx@0.11.0

## 0.5.3

### Patch Changes

- Updated dependencies [d12d6b0]
- Updated dependencies [d12d6b0]
- Updated dependencies [d12d6b0]
  - @consultchimps/xlsx@0.10.0

## 0.5.2

### Patch Changes

- Updated dependencies [128a310]
  - @consultchimps/xlsx@0.9.2

## 0.5.1

### Patch Changes

- Updated dependencies [727239a]
  - @consultchimps/xlsx@0.9.1

## 0.5.0

### Minor Changes

- ea3d302: Add byte-level entry points so Excel and PowerPoint operations can
  run without a filesystem, such as in a browser.

  `@consultchimps/xlsx/bytes` exports `splitWorkbookBytes`,
  `planSplitWorkbookBytes`, `mergeWorkbooksBytes`, and
  `readWorksheetRecordsBytes`. The split accepts the worksheet, Excel Table,
  named range, header row, blank-value, and workbook-preserving options of the
  path-based split, including the formatting-preserving Excel Table split; the
  merge keeps every worksheet's cells and formatting, resolves tab name
  collisions, and reports hidden source worksheets exactly as the path-based
  merge does.

  `@consultchimps/pptx/bytes` exports `populatePresentationBytes`,
  `planPopulatePresentationBytes`, and `inspectPresentationBytes`. The
  population reads its records either from an in-memory array or from workbook
  bytes, and the new `PPTX_INVALID_DATA_SOURCE` error reports a call that
  supplies both or neither.

  Byte operations take named in-memory inputs, return the produced bytes with
  the same structured result the path-based operations report, sanitize every
  output name into a portable filename, support progress reporting and
  cancellation, and produce nothing when cancelled.

  Generated workbooks and presentations are now byte-identical for identical
  inputs: rewritten Open XML parts carry a fixed timestamp instead of the
  current time, and packages no longer gain directory entries the source package
  did not have. This also fixes clock-dependent bytes in the path-based
  worksheet merge, the workbook-preserving split, and PowerPoint population.

### Patch Changes

- Updated dependencies [ea3d302]
  - @consultchimps/xlsx@0.9.0

## 0.4.4

### Patch Changes

- Updated dependencies [6c62d2e]
  - @consultchimps/xlsx@0.8.0

## 0.4.3

### Patch Changes

- Updated dependencies [1f759eb]
  - @consultchimps/core@0.5.0
  - @consultchimps/xlsx@0.7.1
  - @consultchimps/files@0.3.2

## 0.4.2

### Patch Changes

- Updated dependencies [fa3b316]
- Updated dependencies [8b49943]
  - @consultchimps/xlsx@0.7.0

## 0.4.1

### Patch Changes

- Updated dependencies [6564e24]
  - @consultchimps/core@0.4.0
  - @consultchimps/files@0.3.1
  - @consultchimps/xlsx@0.6.1

## 0.4.0

### Minor Changes

- 5c08c75: Publish typed error-code registries (`FILES_ERRORS`, `PDF_ERRORS`,
  `XLSX_ERRORS`, `PPTX_ERRORS`, with matching `*ErrorCode` unions) so consumers
  can match expected failures without string literals, and make
  `OperationResult` and `OperationPlan` generic over each operation's metric
  names so metric renames become compile-time errors. All runtime values and
  error codes are unchanged; the generics default to `string`, so existing
  consumers keep compiling.

### Patch Changes

- Updated dependencies [5c08c75]
- Updated dependencies [5c08c75]
  - @consultchimps/xlsx@0.6.0
  - @consultchimps/core@0.3.0
  - @consultchimps/files@0.3.0

## 0.3.0

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
  - @consultchimps/xlsx@0.5.0

## 0.2.0

### Minor Changes

- 6870be2: Add local PowerPoint template inspection and text population from
  selected Excel worksheet records, with formatting-preserving slide cloning,
  complete pre-write validation, safe overwrite handling, and detailed CLI
  guidance. Also make absolute file-glob discovery portable on Windows.
- 2c4065f: Support PowerPoint placeholders split across adjacent text runs and
  default PowerPoint population to the first template slide and first worksheet
  when those selections are omitted.

### Patch Changes

- Updated dependencies [6870be2]
- Updated dependencies [2c4065f]
  - @consultchimps/xlsx@0.4.0
  - @consultchimps/files@0.1.1
