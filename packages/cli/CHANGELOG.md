# consultchimps

## 0.10.0

### Minor Changes

- 0797532: Add `consultchimps sheets inspect <workbook>`, which describes an
  Excel workbook without creating or changing anything. The report lists each
  described worksheet with its visibility, used range, resolved header row, and
  data row count, every header with a few of the values stored beneath it, and
  the Excel Tables and named ranges those worksheets contain, followed by the
  usual plain-language explanation of the result. `--sheet`, `--header-row`, and
  `--hidden` select what is described, and `--samples` sets how many sample
  values each column reports, from 0 to 5. Under `--json` the envelope carries
  the whole inspection outcome: the description beside the operation result.
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

### Patch Changes

- Updated dependencies [846a8bd]
  - @consultchimps/files@0.4.0
  - @consultchimps/messages@0.5.0
  - @consultchimps/xlsx@0.15.0
  - @consultchimps/pdf@0.4.3
  - @consultchimps/pptx@0.6.3

## 0.9.3

### Patch Changes

- e716dc7: Name a preserved split's outputs after the workbook they actually
  are. Splitting a macro-enabled workbook while naming an Excel Table kept the
  whole source package, macro project included, but named every output `.xlsx`
  and reported the ordinary workbook media type, so the file's contents and its
  name disagreed and Excel opened it with a corruption warning. Those outputs
  are now named `.xlsm` and carry the macro-enabled media type, exactly as the
  all-worksheet split has always done, and a package whose declared type
  contradicts the name it arrived under is refused with
  `XLSX_SPLIT_PACKAGE_TYPE_MISMATCH` before anything is written, on the preview
  as well as the run.

  A split that rebuilds instead of preserving (`preserveWorkbook: false`, a
  named worksheet, or a named range) is unchanged: it writes a fresh ordinary
  package from the rows it kept, carries no macro project, and is still `.xlsx`
  whatever the source was called.

  The CLI carries the same correction:
  `consultchimps sheets split <workbook.xlsm> --table <name>` now writes `.xlsm`
  files and reports them with the macro-enabled media type, and refuses a
  workbook whose package contradicts its name.

- 91e298d: Correct two over-broad claims in the packaged CLI.
  `sheets split --help` said it copies "the complete workbook" once per value;
  it now says "the whole workbook" and states that pivot tables and their caches
  are removed and reported as a warning, which is what the operation has always
  done. The README said the CLI requires Node.js 24 while `engines.node`
  declares `>=22.0.0`; it now says Node.js 22 or later. The README's command
  overview also lists `pptx inspect-template`, which it had never mentioned, and
  points at the CLI reference rather than trying to be a full catalogue.
- Updated dependencies [e716dc7]
- Updated dependencies [76ec2b3]
  - @consultchimps/xlsx@0.14.0
  - @consultchimps/pptx@0.6.2

## 0.9.2

### Patch Changes

- 32973f7: Declare support for Node.js 22 and later (`engines.node: ">=22.0.0"`
  instead of `24.x`), so the toolkit installs and runs in environments that ship
  the previous LTS line. CI now validates the runtime on Node 22.16, the latest
  22, and 26 alongside the full Node 24 verification. The CLI additionally ships
  a standalone `consultchimps.mjs` bundle on each GitHub release: one file that
  runs with `node consultchimps.mjs` and needs no npm access at all.
- Updated dependencies [cef85f7]
- Updated dependencies [32973f7]
- Updated dependencies [cef85f7]
  - @consultchimps/xlsx@0.13.0
  - @consultchimps/core@0.5.1
  - @consultchimps/files@0.3.3
  - @consultchimps/messages@0.4.0
  - @consultchimps/pdf@0.4.2
  - @consultchimps/pptx@0.6.1

## 0.9.1

### Patch Changes

- Updated dependencies [6c683ec]
- Updated dependencies [aabc58b]
- Updated dependencies [c7ccb8d]
- Updated dependencies [cc4d06c]
- Updated dependencies [cc4d06c]
  - @consultchimps/xlsx@0.12.0
  - @consultchimps/messages@0.3.0
  - @consultchimps/pptx@0.6.0

## 0.9.0

### Minor Changes

- a969491: Show live progress on stderr during long-running CLI commands.
  `sheets consolidate`, `sheets merge`, `sheets split`, `pdf split`,
  `pdf merge`, and `pptx populate` now report the current stage and item count
  while they work (for example `Reading workbooks 3/14: report.xlsx`), so large
  jobs no longer sit silent until the final report. In an interactive terminal
  the line updates in place; when output is redirected, plain lines are printed
  instead. Progress goes only to stderr, never interleaves with the final
  report, and is fully suppressed under `--json`.

  `mergeWorkbooks` in `@consultchimps/xlsx` now accepts the standard operation
  controls (`onProgress`, `signal`), emitting `merging-inputs` events per input
  workbook and a final `writing-output` event, matching `consolidateWorkbooks`
  and the in-memory `mergeWorkbooksBytes`.

- 98c77a7: Consolidation can now match columns whose headers differ only in
  case, spacing, or punctuation. Different systems often export the same schema
  with different header conventions - "Failed Checks" in one file,
  "Failed_Checks" in another, "Reviewer: Lead Contact" versus
  "Reviewer_Lead_Contact" - and until now each spelling became its own
  mostly-empty output column.

  Opt in with `--normalize-headers` on `consultchimps sheets consolidate`, or
  `normalizeHeaders: true` on `consolidateWorkbooks` and `unionTables`. Matching
  ignores case and treats any run of spaces or punctuation as one separator; the
  first spelling seen names the output column. The default behaviour is
  unchanged. The tabular package also exports the new `normalizedColumnKey`
  helper behind this matching.

### Patch Changes

- c44f6e6: Clarify the difference between `sheets merge` and
  `sheets consolidate` in CLI help. Each command's description now states the
  shape of the result up front (merge keeps every worksheet as its own tab;
  consolidate stacks all rows into one combined sheet, matching columns by
  header), and each command's help now points to the other command for when that
  result is what you want.
- Updated dependencies [a969491]
- Updated dependencies [98c77a7]
- Updated dependencies [d1f8524]
  - @consultchimps/xlsx@0.11.0
  - @consultchimps/pptx@0.5.4

## 0.8.4

### Patch Changes

- Updated dependencies [d12d6b0]
- Updated dependencies [d12d6b0]
- Updated dependencies [d12d6b0]
  - @consultchimps/xlsx@0.10.0
  - @consultchimps/pptx@0.5.3

## 0.8.3

### Patch Changes

- 128a310: Compact retained plain worksheet rows after splitting so deleted rows
  do not remain as visible gaps in generated workbooks.
- Updated dependencies [128a310]
  - @consultchimps/xlsx@0.9.2
  - @consultchimps/pptx@0.5.2

## 0.8.2

### Patch Changes

- 727239a: Physically remove unmatched worksheet rows during preserved Excel
  splits instead of leaving empty row shells behind.
- Updated dependencies [727239a]
  - @consultchimps/xlsx@0.9.1
  - @consultchimps/pptx@0.5.1

## 0.8.1

### Patch Changes

- Updated dependencies [ea3d302]
  - @consultchimps/pptx@0.5.0
  - @consultchimps/xlsx@0.9.0

## 0.8.0

### Minor Changes

- 6c62d2e: Split complete Excel workbooks by normalized column values collected
  across all worksheets, preserving workbook formatting, formulas or cached
  values, tables, VBA content, and sheets that do not contain the selected
  column.

### Patch Changes

- Updated dependencies [6c62d2e]
  - @consultchimps/xlsx@0.8.0
  - @consultchimps/messages@0.2.4
  - @consultchimps/pptx@0.4.4

## 0.7.0

### Minor Changes

- 7199fb8: Give `--json` a single-line result envelope on stdout. A successful
  command now prints `{"ok":true,"result":...}` and a failing command prints
  `{"ok":false,"error":{"message":...,"code":...}}` while keeping its nonzero
  exit code, so an automation can branch on `ok` and read the stable error code
  without parsing human text. Usage errors such as an unknown option or a
  missing required option are reported through the same envelope with the code
  `CLI_USAGE`, instead of leaving stdout empty. `--help` and `--version` are
  unchanged.

  This changes the shape of `--json` output. Commands previously printed the
  operation result pretty-printed and unwrapped, and a failure wrote a prose
  line to stderr with nothing machine-readable on stdout. Consumers that read
  `--json` should now read the value under `result`.

### Patch Changes

- Updated dependencies [1f759eb]
  - @consultchimps/core@0.5.0
  - @consultchimps/pdf@0.4.1
  - @consultchimps/xlsx@0.7.1
  - @consultchimps/files@0.3.2
  - @consultchimps/messages@0.2.3
  - @consultchimps/pptx@0.4.3

## 0.6.0

### Minor Changes

- fa3b316: Add a values-only option to every Excel operation so formulas can be
  replaced with their stored results without removing workbook formatting.
- 8b49943: Add `sheets merge` to copy formatted worksheets from multiple Excel
  workbooks into one workbook while recording hidden-sheet status.

### Patch Changes

- Updated dependencies [fa3b316]
- Updated dependencies [8b49943]
  - @consultchimps/xlsx@0.7.0
  - @consultchimps/messages@0.2.2
  - @consultchimps/pptx@0.4.2

## 0.5.2

### Patch Changes

- Updated dependencies [6564e24]
  - @consultchimps/core@0.4.0
  - @consultchimps/pdf@0.4.0
  - @consultchimps/files@0.3.1
  - @consultchimps/messages@0.2.1
  - @consultchimps/pptx@0.4.1
  - @consultchimps/xlsx@0.6.1

## 0.5.1

### Patch Changes

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

- Updated dependencies [00740ec]
  - @consultchimps/messages@0.2.0

## 0.5.0

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

### Patch Changes

- 5c08c75: Publish typed error-code registries (`FILES_ERRORS`, `PDF_ERRORS`,
  `XLSX_ERRORS`, `PPTX_ERRORS`, with matching `*ErrorCode` unions) so consumers
  can match expected failures without string literals, and make
  `OperationResult` and `OperationPlan` generic over each operation's metric
  names so metric renames become compile-time errors. All runtime values and
  error codes are unchanged; the generics default to `string`, so existing
  consumers keep compiling.
- Updated dependencies [849bf37]
- Updated dependencies [5c08c75]
- Updated dependencies [5c08c75]
  - @consultchimps/pdf@0.3.0
  - @consultchimps/xlsx@0.6.0
  - @consultchimps/core@0.3.0
  - @consultchimps/files@0.3.0
  - @consultchimps/pptx@0.4.0
  - @consultchimps/messages@0.1.1

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
