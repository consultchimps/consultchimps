# consultchimps

## 0.4.0

### Minor Changes

- 139b545: Add local PowerPoint template inspection and text population from
  selected Excel worksheet records, with formatting-preserving slide cloning,
  complete pre-write validation, safe overwrite handling, and detailed CLI
  guidance. Also make absolute file-glob discovery portable on Windows.

### Patch Changes

- 6d9a7fb: Make every CLI tool explain successful results and recoverable errors
  in detailed, non-technical language while preserving compact JSON output for
  automation.
- 31ff5a0: Make output options easier to discover by adding practical examples
  to every CLI help level.
- Updated dependencies [139b545]
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
