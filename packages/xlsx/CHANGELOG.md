# @consultchimps/xlsx

## 0.6.0

### Minor Changes

- 5c08c75: Excel Table splits now preserve the complete source workbook by
  default, so every output opens exactly like the prepared original (zoom,
  cursor position, cover sheets, styles, and content outside the table all carry
  over). Pass `preserveWorkbook: false` or the new `--no-preserve-workbook` flag
  for the previous compact data-only outputs. Behavior change for existing table
  splits that relied on the compact default; plain worksheet splits are
  unchanged.

  Preserved splits now refuse to relocate cells whose formulas depend on their
  position (A1-style references, shared, or array formulas) with a stable
  `XLSX_SPLIT_PRESERVE_FORMULA` error instead of silently producing formulas
  that point at the wrong rows; structured table references such as `[@Amount]`
  remain fully supported.

  Workbook named ranges are now supported as a data source: select one with the
  new `range` option or `--range <name>` flag, inspect them with the new
  `readWorkbookNamedRanges` export, and prefer sources in the order Excel Table,
  named range, then full worksheet range. Named-range splits always produce
  compact outputs and reject `headerRow` and `preserveWorkbook`.

- 5c08c75: Publish typed error-code registries (`FILES_ERRORS`, `PDF_ERRORS`,
  `XLSX_ERRORS`, `PPTX_ERRORS`, with matching `*ErrorCode` unions) so consumers
  can match expected failures without string literals, and make
  `OperationResult` and `OperationPlan` generic over each operation's metric
  names so metric renames become compile-time errors. All runtime values and
  error codes are unchanged; the generics default to `string`, so existing
  consumers keep compiling.

### Patch Changes

- Updated dependencies [5c08c75]
  - @consultchimps/core@0.3.0
  - @consultchimps/files@0.3.0
  - @consultchimps/tabular@0.2.2

## 0.5.0

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
  - @consultchimps/tabular@0.2.1

## 0.4.0

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
  - @consultchimps/files@0.1.1

## 0.3.0

### Minor Changes

- 0b3f5cd: Add an opt-in Excel Table split mode that preserves the complete
  source workbook, including its worksheets, formatting, layout, and content
  outside the selected table.

## 0.2.0

### Minor Changes

- 7958e80: Add reusable table grouping and an Excel split-by-column API and CLI
  command with deterministic filenames, blank-row controls, and safe output
  preflight.
- a019777: Allow spreadsheet splitting to select a named Excel Table, excluding
  totals and all worksheet cells outside the table range.

### Patch Changes

- Updated dependencies [7958e80]
  - @consultchimps/tabular@0.2.0
