# @consultchimps/messages

## 0.7.0

### Minor Changes

- 4bc99d2: Raise the supported Node.js floor from 22.0.0 to 22.14.0. Continuous
  integration now tests that exact version: pnpm 11 requires 22.13, and
  better-sqlite3 13 crashes on the 22.13 patch line, so 22.14.0 is the lowest
  release on which every published package works.
- 6606d1e: Group staged imports under `db import prepare`, `inspect`, `update`,
  and `apply`. Use `db import run` for one-step execution, `--profile` for
  reusable settings, and `--batch` for a saved review. Rename the library
  contracts to import profiles and batches while preserving the current stored
  database and batch formats.

  Page saved import routes independently from row previews and expose applied
  capture membership in import inspection. Share destination naming policies
  between callers while preserving their Record ID prefix conventions. Count
  shared captures once in review totals while retaining each source binding.

  Report whether database writes committed or returned an unchanged result. Keep
  completed write results available after checkpoint or summary-refresh
  failures. Add retryable resource ownership through `OwnedResources` in the
  core package.

  Create and reopen persistent SQLite and DuckDB databases, prepare workbook
  imports for review, and retain source captures and separately recorded
  batches. Add bounded workbook reading and file access for large imports.

  Replace the database spike's synchronous in-memory API with asynchronous
  persistent operations. Applications must migrate to the new database and
  runtime entry points. Saved source files are not automatically rewritten.

  Add the `consultchimps db` command group. Portable CLI distribution now
  includes native engine assets in an archive for its operating system,
  architecture, and Node major, rather than a single JavaScript file.

  Bind prepared artifact format 3 reviews to their metadata fingerprint and
  capture row checksums. Verify consumed row values before committing an import.
  Version 1 and 2 spike batches must be regenerated from their source files;
  existing working database files keep their format.

  Accept leading-plus integer tokens consistently during inference and import.
  Keep database and import-batch replacement blocked after failed closure until
  a close retry succeeds.

  Reject ill-formed Unicode in schema identifiers and Record ID components.
  Reject added, missing, or reordered internal metadata columns when opening
  working databases and saved import batches.

  Preserve supplementary Unicode characters in inferred table names and Record
  ID prefixes. Keep streaming workbook cleanup retryable after a scratch-file
  close failure, and report failed browser candidate cleanup with recovery
  locations.

  Reject automatic imports with no selected workbook regions before recording an
  import. Report retained native staging files and outputs published before a
  reopen failure. Keep scratch cleanup retryable and block native replacement
  after failed registration cleanup until the unresolved handle closes. Report
  successful publication when removing the staging filename fails, and retain
  private files when initialization leaves an unclosed engine connection. Reject
  SQLite triggers on managed or internal tables before managed writes so they
  cannot silently suppress imported rows or alter retry receipts.

  Revalidate schema plans before applying changes and reject undeclared reader
  cells before capture. Preserve cleanup failures and recovery paths when a
  native open or one-step CLI import fails.

  Validate internal table types, nullability, and key constraints when opening
  databases and saved batches. Delay CLI database results until resource cleanup
  finishes, preserving valid JSON and recovery details after a committed
  operation. Check cancellation during browser replacement copies and preserve
  failed browser open cleanup so live owners continue to block replacement.

  Reject repeated workbook-property declarations and duplicate import source or
  selection keys before capture. Reject stored application counts that disagree
  with their capture before reporting reuse or recording new retry receipts.

  Reject browser automatic imports when no workbook regions are selected. Check
  metadata allocation counters when opening a database and row allocation before
  new table applications. Report structured corruption errors for counter
  conflicts and pending Record ID collisions while preserving valid counter gaps
  and the already-applied retry path.

  Reject import receipt replay when the approved batch belongs to another
  database, including a copy converted to a different format.

  Parse workbook row references in linear time so malformed references cannot
  trigger excessive regular-expression backtracking.

  Detect DuckDB view changes before applying reviewed schema or import batches.
  Validate saved import receipt totals against application metadata before
  returning a successful retry result.

  Read the workbook date system only from its SpreadsheetML property element so
  extension metadata cannot shift imported dates.

  Reject saved import batches containing rows from an unfinished capture when
  reopening them, without deleting the batch or scanning row values.

  Ignore foreign or misplaced workbook sheet and named-range declarations.
  Reject retry receipts that substitute applications from unrelated captures or
  tables while retaining reuse of applications from earlier requests.

  Apply namespace and ancestry checks to worksheet values, shared strings,
  styles, and table references. Advance the workbook reader version so new
  preparations read files again when only an older reader capture exists.
  Existing imported rows and capture history are not rewritten automatically.

  Reject unsupported schema and import-profile versions from JavaScript callers
  before publishing database or prepared-batch outputs.

  Reject unsupported XML document roots instead of reporting an empty workbook.
  Verify batch capture memberships before returning an import retry receipt.

  Reject worksheet aliases and require supported OOXML relationship role URIs.
  Preserve orphaned browser DuckDB recovery files by refusing publication over
  incomplete storage.

  Report committed batch updates when preparation checkpointing fails.
  Distinguish caller-owned batches from private staging discarded before
  publication. Preserve DuckDB appender operation and cleanup failures together,
  and retry only the remaining native resources after a partially failed
  database close.

  Preserve the original workbook initialization error when lazy opening fails.
  Closing a source with no acquired stream no longer creates a false cleanup
  failure or blocks later preparations.

  Preserve transaction and rollback failures together in SQLite and DuckDB,
  across native and browser runtimes. Block further operations on an unresolved
  connection while leaving cleanup available, and retain the original error when
  rollback succeeds.

### Patch Changes

- 78ff6a6: Skip title rows and spacer columns when reading worksheets. The
  header row is now the first row holding more than half as many values as the
  fullest of the ten populated rows below it, or at most one value fewer, so a
  report title, a merged banner, or a "Prepared by" line above a table four or
  more columns wide no longer becomes the first column name. A row is skipped
  only on evidence that it cannot be the header of the block below it, a blank
  row between the two or every value on it sitting in a cell merged across
  columns, and never when the skipped rows form a table of their own (two
  adjacent populated rows, neither a banner); otherwise the first row holding a
  value stays the header, as it always was, so no row of a table is ever lost to
  the detection. The inspection's `headerColumns` metric now counts the kept
  columns. A column that holds nothing in the header row or in any row under it
  is a spacer and never reaches the output; a column with values under a blank
  header is kept and named `column_N` by its position among the columns that
  were kept. One rule decides this for the consolidation, the worksheet readers,
  the worksheet-records reader, and the inspection, so `sheets inspect` reports
  the header row and columns a consolidation will use, spelled the same way
  whatever the header cell's type; an error cell in a header row now names its
  column `#REF!` rather than Excel's internal code for it. A split looks for its
  column on the detected header row before searching the whole worksheet, so a
  title line that repeats the column's name no longer captures the split. A
  declared `headerRow` is never second-guessed. Consolidation results carry two
  new metrics, `skippedTitleRows` and `skippedSpacerColumns`, and every
  worksheet report carries the same two counts; the plain-language result says
  what was left out.
- Updated dependencies [4bc99d2]
- Updated dependencies [6606d1e]
  - @consultchimps/core@0.6.0

## 0.6.0

### Minor Changes

- 1eec990: Explain a workbook unprotect result in plain language instead of the
  generic fallback. The new `sheets.unprotect` entry reports how many worksheet
  and workbook-structure protections were removed, including the case where
  there was nothing to remove, and a macro-enabled `.xlsm` output is now named
  as an Excel workbook rather than a bare file.

## 0.5.0

### Minor Changes

- 846a8bd: Consolidation can now fold columns that are named differently into
  one column each, using the versioned JSON column mapping document.

  `sheets consolidate` gains `--map <file>`, which applies a mapping before the
  rows are stacked, and `--suggest-map <file>`, which writes a draft mapping
  built from the headers that were read and still writes the consolidated
  workbook. A draft is never applied for you, it goes through the same
  never-overwrite rule as any other output, and the two options cannot be
  combined in one run.

  A column no mapping entry claims keeps its own name and is reported as a
  warning. Two columns of one worksheet folding into one canonical column stop
  the run before anything is written. A declared date coercion reads text: a
  column holding a number, or a value Excel already stores as a date, is refused
  by name rather than read as a date serial, because which day a serial counts
  from belongs to the workbook rather than to the cell.

  `@consultchimps/files` gains `isSameFilesystemPath` and `isPathWithin`, the
  destination-collision checks an operation with more than one output needs: two
  names that differ only in case are one file on Windows and the usual macOS
  volume, and one output can never sit inside another.

  `consolidateWorkbooks` takes `mappingFile` and `suggestMappingOutput`, and
  `consolidateWorkbooksBytes` takes a parsed `mapping` and `suggestMapping`.
  Both return the drafted mapping on the result, and both report two new
  metrics, `unmappedColumns` and `suggestedColumns`. Result explanations name
  the columns a mapping did not claim, point at a drafted mapping among the
  created files, and give mapping failures their own recovery steps.

## 0.4.0

### Minor Changes

- cef85f7: Explain workbook inspection results in plain language. A
  `sheets.inspect` result now renders its own summary and next steps: counts of
  worksheets, columns, data rows, Excel Tables, and named ranges, with wording
  that never points at output files, because an inspection creates none. Its
  metrics read as readable labels rather than internal names.

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
  inspection metrics (malformed placeholder locations, placeholders outside a
  supported text shape, and placeholders split across text runs) instead of
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
