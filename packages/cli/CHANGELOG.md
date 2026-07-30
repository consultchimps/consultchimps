# consultchimps

## 0.4.1

### Patch Changes

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
- c78b35e: Extract the CLI's plain-language result and error rendering into the
  new `@consultchimps/messages` package so desktop and web interfaces can reuse
  the same explanations. The CLI output is unchanged apart from a new recovery
  explanation for cancelled operations.
- Updated dependencies [c78b35e]
- Updated dependencies [c78b35e]
  - @consultchimps/core@0.2.0
  - @consultchimps/files@0.2.0
  - @consultchimps/xlsx@0.5.0
  - @consultchimps/pptx@0.3.0
  - @consultchimps/pdf@0.2.0
  - @consultchimps/messages@0.1.0

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

- 6d9a7fb: Make every CLI tool explain successful results and recoverable errors
  in detailed, non-technical language while preserving compact JSON output for
  automation.
- 31ff5a0: Make output options easier to discover by adding practical examples
  to every CLI help level.
- Updated dependencies [6870be2]
- Updated dependencies [2c4065f]
  - @consultchimps/xlsx@0.4.0
  - @consultchimps/pptx@0.2.0
  - @consultchimps/files@0.1.1
  - @consultchimps/pdf@0.1.1

## 0.3.0

### Minor Changes

- 0b3f5cd: Add an opt-in Excel Table split mode that preserves the complete
  source workbook, including its worksheets, formatting, layout, and content
  outside the selected table.

### Patch Changes

- Updated dependencies [0b3f5cd]
  - @consultchimps/xlsx@0.3.0

## 0.2.0

### Minor Changes

- 7958e80: Add reusable table grouping and an Excel split-by-column API and CLI
  command with deterministic filenames, blank-row controls, and safe output
  preflight.
- a019777: Allow spreadsheet splitting to select a named Excel Table, excluding
  totals and all worksheet cells outside the table range.

### Patch Changes

- Updated dependencies [7958e80]
- Updated dependencies [a019777]
  - @consultchimps/xlsx@0.2.0
