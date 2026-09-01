# @consultchimps/xlsx

## 0.15.0

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

### Patch Changes

- Updated dependencies [846a8bd]
  - @consultchimps/files@0.4.0

## 0.14.0

### Minor Changes

- 76ec2b3: Export the conformance contract: `CONTRACT`, `TRACKED_STRUCTURES`,
  `OPERATIONS`, the `UNDECIDED_*` records, and their types now ship from the
  package root. The table states, per workbook structure and operation, whether
  the package preserves it, rewrites it so it stays valid, removes it with a
  warning, or refuses, using the same table the corpus tests enforce, so a
  caller can tell people what an operation will do before running it, and the
  documentation site can generate its preservation matrix instead of restating
  it in prose.

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

## 0.13.0

### Minor Changes

- cef85f7: Add workbook inspection: `describeWorkbook` and
  `describeWorkbookBytes` report a workbook's structure without creating
  anything: worksheets with their visibility and dimensions, the header row an
  operation would actually use, Excel Tables, named ranges, and up to five
  distinct sample values per column. The outcome pairs that description with a
  structured `sheets.inspect` result carrying counts as metrics, no artifacts,
  and warnings for hidden worksheets left out and worksheets with no header row.

  The bytes surface also gains `readWorkbookExcelTablesBytes` and
  `readWorkbookNamedRangesBytes`, the byte twins of the path-based Excel Table
  and named-range readers, so its `WorkbookExcelTable` and `WorkbookNamedRange`
  type exports now describe values that surface can actually produce.

### Patch Changes

- 32973f7: Declare support for Node.js 22 and later (`engines.node: ">=22.0.0"`
  instead of `24.x`), so the toolkit installs and runs in environments that ship
  the previous LTS line. CI now validates the runtime on Node 22.16, the latest
  22, and 26 alongside the full Node 24 verification. The CLI additionally ships
  a standalone `consultchimps.mjs` bundle on each GitHub release: one file that
  runs with `node consultchimps.mjs` and needs no npm access at all.
- Updated dependencies [32973f7]
- Updated dependencies [7f309f6]
  - @consultchimps/core@0.5.1
  - @consultchimps/files@0.3.3
  - @consultchimps/tabular@0.4.0

## 0.12.0

### Minor Changes

- 6c683ec: Split whole workbooks by default without a filesystem, so the
  in-browser splitter now works the way the command line does.

  `splitWorkbookBytes` and `planSplitWorkbookBytes` previously read a single
  worksheet and wrote compact, data-only workbooks. They now collect the values
  of the chosen column across every worksheet and give back one complete copy of
  the source workbook per value:

  - a worksheet that carries the column keeps its header and only that value's
    rows;
  - a worksheet that does not carry the column is copied through untouched; and
  - sheet order and visibility, formatting, merged cells, conditional
    formatting, data validation, hyperlinks, comments, images, charts, defined
    names, and the macro parts of an `.xlsm` file all survive.

  `.xlsm` workbooks are accepted as input and produce `.xlsm` outputs.

  A workbook whose contents disagree with its file extension is now refused, on
  both the command line and in the browser, with a new
  `XLSX_SPLIT_PACKAGE_TYPE_MISMATCH` error naming which side to correct. An
  ordinary workbook renamed `.xlsm`, or a macro-enabled one renamed `.xlsx`,
  used to be split into files Excel could open with a corruption warning; the
  split now stops before writing anything.

  Strict matching is available, so case, surrounding whitespace, and value type
  can be kept distinct instead of `North`, `north`, and `North ` becoming one
  workbook.

  Results report more of what happened: how many worksheets were filtered, how
  many were copied unchanged, and, for each output workbook, the rows kept and
  removed per worksheet. The values-only, pivot-cache and stale-cached-total
  warnings match the ones the command line reports.

  The single-source split is still available for narrower jobs: name an Excel
  Table, a named range, or a worksheet, or turn off whole-workbook preservation
  for compact data-only outputs.

- c7ccb8d: Consolidate workbooks without a filesystem.
  `@consultchimps/xlsx/bytes` now exports `consolidateWorkbooksBytes`, which
  takes named in-memory workbooks and returns one combined workbook's bytes
  alongside the same structured result the path-based `consolidateWorkbooks`
  reports. It accepts the same `normalizeHeaders`, `addSourceColumns`,
  worksheet-selection, and header-row options, reports progress, and can be
  cancelled, and it produces byte-identical output to the path-based operation
  for the same workbooks and options, so a browser and the command line agree
  exactly.

### Patch Changes

- aabc58b: Installing this package no longer requires network access to the
  SheetJS CDN. SheetJS was declared as a runtime dependency pointing at a
  tarball URL, so every `npm install` of `@consultchimps/xlsx`, and of the
  `consultchimps` CLI that depends on it, had to reach `cdn.sheetjs.com`.
  Installs failed outright behind registry-only allowlists, corporate proxies,
  private mirrors, and locked-down CI runners.

  SheetJS is now compiled into the published `dist` output instead, so the
  package installs from the npm registry alone. Behaviour, the public API, and
  the generated type declarations are unchanged; the published bundle is
  correspondingly larger. The bundled Apache-2.0 code is attributed in the
  package's `THIRD-PARTY-LICENSES.md`.

## 0.11.0

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

- d1f8524: Write consolidated and rebuilt split workbooks with a shared-strings
  table instead of per-cell strings, matching how Excel stores text. Outputs
  with repetitive text now serialize noticeably smaller: roughly 13-20% on
  synthetic benchmarks, with the biggest gains when repeated values are spread
  far apart in large workbooks.
- Updated dependencies [98c77a7]
  - @consultchimps/tabular@0.3.0

## 0.10.0

### Minor Changes

- d12d6b0: The workbook merge now preserves Excel Tables, defined names,
  conditional formatting, data validation, cell comments, styles and number
  formats at the package level instead of rebuilding each worksheet through a
  spreadsheet library.

  `mergeWorkbooks` and `mergeWorkbooksBytes` keep their signatures, their
  metrics and their error codes. What changed is the engine underneath: the
  first input now seeds the output package and every later input's worksheet
  parts are copied into it with the parts they depend on, so a structure
  survives unless carrying it would be wrong.

  - **Preserved**: merged cells, conditional formatting, data validation,
    hyperlinks, comments and their drawings, Excel Tables including totals rows,
    cell styles and number formats, and shared, array and uncached formulas.
  - **Repaired**: shared-string and style indexes are remapped into one merged
    table per workbook (identical entries collapse); duplicate Excel Table names
    and workbook-scoped defined names take a numeric suffix, with a warning
    listing every rename; formulas that named a renamed worksheet or table -
    structured references included - are rewritten to follow it.
  - **Removed, each with a warning**: pivot tables and their caches, external
    links, and a macro project that cannot travel. A macro project is carried
    only when a single input has one and the output is named `.xlsm`; the byte
    surface now keeps an `.xlsm` output name a caller asks for, and reports the
    macro-enabled media type for it.
  - The calculation chain is dropped without a warning and the merged workbook
    asks Excel to recalculate on open, because a chain is a derived index keyed
    by sheet ids that every transplanted worksheet changes.

- d12d6b0: Fix every reference that pointed at a row a workbook split moved, and
  make split outputs byte-reproducible.

  Splitting a workbook by column removes the rows that belong to other groups
  and closes the gaps they leave. Until now only the rows and their cells were
  renumbered: everything else that described those rows kept its original row
  number, so a delivered workbook could highlight the wrong cells, validate the
  wrong column, link from the wrong row, or double the wrong record.

  All of it now moves with the rows, on both worksheet ranges and Excel Tables:

  - merged cells,
  - conditional-formatting and data-validation ranges,
  - hyperlinks,
  - cell comments, including the drawing anchor that positions the note,
  - formulas, including shared and array formulas and the spans they claim,
  - the worksheet's declared used range.

  The split also writes every output through one deterministic package writer,
  so splitting the same workbook twice now produces byte-identical files rather
  than files whose contents match but whose timestamps do not. Parts the split
  does not touch still travel through byte for byte.

  Public API, error codes, metrics and warning shapes are unchanged.

- d12d6b0: Close the first three Tier-1 gaps in the workbook split.

  - Pivot caches no longer leak other groups' rows into split outputs. A pivot
    cache is a private copy of the source rows that travels inside the package,
    so every pivot table and cache part is now removed from each output,
    together with its relationships, content-type overrides and the workbook's
    `pivotCaches` registry. The result reports how many were removed and warns
    that the pivot has to be rebuilt.
  - Values-mode splits no longer bake aggregates computed over removed rows.
    Before the values conversion runs, the cached result of any formula whose A1
    references reach into rows the group does not receive -- including
    cross-sheet references -- is cleared, so the output shows a reported blank
    cell rather than another group's total presented as this one's.
  - The calculation chain is kept consistent. Entries naming cells a split
    deleted are dropped, entries whose rows moved are renumbered, and an emptied
    chain is removed along with its relationship and content-type override.

## 0.9.2

### Patch Changes

- 128a310: Compact retained plain worksheet rows after splitting so deleted rows
  do not remain as visible gaps in generated workbooks.

## 0.9.1

### Patch Changes

- 727239a: Physically remove unmatched worksheet rows during preserved Excel
  splits instead of leaving empty row shells behind.

## 0.9.0

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

## 0.8.0

### Minor Changes

- 6c62d2e: Split complete Excel workbooks by normalized column values collected
  across all worksheets, preserving workbook formatting, formulas or cached
  values, tables, VBA content, and sheets that do not contain the selected
  column.

## 0.7.1

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
  - @consultchimps/tabular@0.2.4

## 0.7.0

### Minor Changes

- fa3b316: Add a values-only option to every Excel operation so formulas can be
  replaced with their stored results without removing workbook formatting.
- 8b49943: Add `sheets merge` to copy formatted worksheets from multiple Excel
  workbooks into one workbook while recording hidden-sheet status.

## 0.6.1

### Patch Changes

- Updated dependencies [6564e24]
  - @consultchimps/core@0.4.0
  - @consultchimps/files@0.3.1
  - @consultchimps/tabular@0.2.3

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
