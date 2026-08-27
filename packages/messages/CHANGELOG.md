# @consultchimps/messages

## 0.4.0

### Minor Changes

- cef85f7: Explain workbook inspection results in plain language. A
  `sheets.inspect` result now renders its own summary and next steps — counts of
  worksheets, columns, data rows, Excel Tables, and named ranges, with wording
  that never points at output files, because an inspection creates none — and
  its metrics read as readable labels rather than internal names.

### Patch Changes

- 32973f7: Declare support for Node.js 22 and later (`engines.node: ">=22.0.0"`
  instead of `24.x`), so the toolkit installs and runs in environments that ship
  the previous LTS line. CI now validates the runtime on Node 22.16, the latest
  22, and 26 alongside the full Node 24 verification. The CLI additionally ships
  a standalone `consultchimps.mjs` bundle on each GitHub release: one file that
  runs with `node consultchimps.mjs` and needs no npm access at all.
- Updated dependencies [32973f7]
  - @consultchimps/core@0.5.1

## 0.3.0

### Minor Changes

- cc4d06c: Explain a PowerPoint template inspection in plain language.
  `formatHumanResult` now recognises the `pptx.inspect-template` operation: it
  says what the slide contains, states that nothing was created or changed
  rather than pointing at output files that do not exist, and labels the
  inspection metrics — malformed placeholder locations, placeholders outside a
  supported text shape, and placeholders split across text runs — instead of
  printing their internal names.

## 0.2.4

### Patch Changes

- 6c62d2e: Split complete Excel workbooks by normalized column values collected
  across all worksheets, preserving workbook formatting, formulas or cached
  values, tables, VBA content, and sheets that do not contain the selected
  column.

## 0.2.3

### Patch Changes

- Updated dependencies [1f759eb]
  - @consultchimps/core@0.5.0

## 0.2.2

### Patch Changes

- 8b49943: Add `sheets merge` to copy formatted worksheets from multiple Excel
  workbooks into one workbook while recording hidden-sheet status.

## 0.2.1

### Patch Changes

- Updated dependencies [6564e24]
  - @consultchimps/core@0.4.0

## 0.2.0

### Minor Changes

- 00740ec: Let every interface reuse the plain-language explanations in its own
  words. `formatHumanResult` and `formatHumanError` now accept an optional
  `{ vocabulary }` option typed as the new exported `MessageVocabulary`, which
  holds the interface-specific phrases: how to retry with overwriting enabled,
  where to find the reference or examples, how to inspect a PowerPoint template
  first, how to point at the created files, and the word for a unit of work.

  Two vocabularies ship with the package. `GENERIC_VOCABULARY` is the new
  default and never names a flag, an executable, or a terminal, so a desktop or
  browser interface can show the guidance unchanged. `CLI_VOCABULARY` reproduces
  the command-line wording, including `--force`, `--help`, and
  `consultchimps pptx inspect-template`.

  Both functions keep their existing signatures, so current callers still
  compile. Library callers that do not pass a vocabulary now receive the neutral
  wording instead of command-line instructions. The `consultchimps` CLI passes
  `CLI_VOCABULARY`, so its output is byte-for-byte unchanged.

## 0.1.1

### Patch Changes

- 5c08c75: Publish typed error-code registries (`FILES_ERRORS`, `PDF_ERRORS`,
  `XLSX_ERRORS`, `PPTX_ERRORS`, with matching `*ErrorCode` unions) so consumers
  can match expected failures without string literals, and make
  `OperationResult` and `OperationPlan` generic over each operation's metric
  names so metric renames become compile-time errors. All runtime values and
  error codes are unchanged; the generics default to `string`, so existing
  consumers keep compiling.
- Updated dependencies [5c08c75]
  - @consultchimps/core@0.3.0

## 0.1.0

### Minor Changes

- c78b35e: Extract the CLI's plain-language result and error rendering into the
  new `@consultchimps/messages` package so desktop and web interfaces can reuse
  the same explanations. The CLI output is unchanged apart from a new recovery
  explanation for cancelled operations.

### Patch Changes

- Updated dependencies [c78b35e]
  - @consultchimps/core@0.2.0
