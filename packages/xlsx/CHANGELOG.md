# @consultchimps/xlsx

## 0.18.0

### Minor Changes

- 04dd6de: Import a table of values into a local database.

  `@consultchimps/db` gains the import half of the workspace, as runtime-neutral
  public API: `inferColumnTypes` decides each column's type from its own values,
  `importTable` and `importTables` create the tables and fill them with
  generated Record IDs, `importedTableSchema` shows what an import would create
  without creating it, and `suggestTableName` and `suggestRecordIdPrefix` derive
  the starting points a person edits before committing. `parseCsvTable` reads
  comma-separated text (RFC 4180 quoting, CRLF, LF, and CR line endings, a
  leading byte-order mark) as a `Table` of text values, so the type of a column
  is decided in one place rather than guessed by a spreadsheet engine while it
  parses.

  Inference is conservative and never guesses: a column takes `boolean`, `date`,
  `integer`, or `real` only when every value that is not blank fits it, and
  stays `text` otherwise. Both numeric rules run one round-trip test, so a value
  is read as a number only when it reads back as the number it wrote: a padded
  reference code such as `007`, an exponent form, a whole number past the range
  stored exactly, and a decimal carrying more digits than a number can hold all
  stay text. A timestamp is judged as one value, so hour 24 is accepted only as
  the end of a day, second 60 only as a leap second at 23:59, and an offset only
  within its real range. An import of several tables is one unit: names, Record
  ID prefixes, and column names are checked first, then every table is created
  and filled inside a single transaction, so a failure part way through leaves
  the database exactly as it was.

  A `date` column now holds ISO 8601 text and nothing else. `sqlValueFromCell`,
  the single point every write goes through, judges a date with `isIsoDateText`,
  the same rule the import's inference uses to decide a column is a date, and
  refuses anything else with `DB_INVALID_DATE`; reading back a stored value that
  is not one reports damage. Surrounding spaces are removed before the value is
  judged and stored, and a value of nothing but spaces is stored as null, so a
  column that claims a spelling holds it and an empty cell is findable as empty.
  This is a behaviour change: a date column that previously accepted any text
  now accepts only a date.

  A source whose columns repeat a name exactly is refused with
  `DB_IMPORT_DUPLICATE_SOURCE_COLUMN`. A `Table` row is an object keyed by
  column name, so the second `Amount` in `["Amount", "Amount"]` is a column no
  row can answer: the remap read the one property twice and stored the same
  value under two names, manufacturing a duplicate the source never had. Names
  differing only in case, or in what the identifier fold ignores, are two
  properties carrying two values and are still numbered and reported, not
  refused. Every reader here makes its headers unique, so this reaches only a
  `Table` a caller built.

  `importedTableSchema` now refuses what the import refuses. The column, table
  name and Record ID checks ran beside the shared plan rather than inside it, so
  a source with no columns, or one whose only column is the Record ID, previewed
  as a schema with no columns while the import raised `DB_IMPORT_NO_COLUMNS` and
  created nothing, and a preview would describe a table under a name the engine
  will not take. They live in the plan now; only the checks that need the
  workspace, such as a name it already holds, still belong to the import.

  `suggestTableName` offers a cleaned name at the full identifier limit. It used
  to leave room for the `table_` prefix in every suggestion, including the ones
  nothing would ever prefix, so a name between 195 and 200 characters came back
  shortened for a prefix it was not going to get, and two files differing only
  in those last characters were handed the same name to resolve by hand. The
  room is now taken in the branch that adds the prefix, from the one name that
  needs it, and the result fits by construction either way.

  A column name is derived in the same step that numbers repeated ones, so it is
  inside the identifier limit by construction: a header too long to be a name is
  shortened at a whole-character boundary with room left for the number a
  duplicate may need, and `ImportedTable.renamedColumns` reports every header
  that ended up under a different name. Shortening after numbering, in a
  different layer, refused a whole import when two copies of a 200-character
  header became a name of 202.

  Also new: `Database.countRecords` counts a table's records in the engine
  rather than by reading its rows, `assertRecordIdConfig` and
  `MAX_RECORD_ID_PADDING` make the Record ID rules callable,
  `MAX_IDENTIFIER_LENGTH` and `truncateIdentifier` state the identifier limit
  and shorten a name to fit it without ever cutting a character in half, and
  `@consultchimps/db/schema` publishes the schema model on its own, with no
  database engine behind it, so a browser page can fold an identifier without
  downloading WebAssembly to do it. `assertSafeIdentifier` now also refuses a
  name that is not well-formed text: an incomplete character passed every other
  check and was stored as U+FFFD, so the name in the database was not the name
  that was asked for.

  `@consultchimps/xlsx` gains `readWorkbookTablesBytes`, the byte twin of
  `readWorkbookTables`: every visible worksheet that holds data as a `Table`,
  with the same `sheets`, `headerRow`, and `includeHiddenSheets` options. The
  byte surface had readers for Excel Tables and named ranges but none for the
  worksheets themselves.

  `readWorkbookWorksheets` and `readWorkbookWorksheetsBytes` join it: every
  selected worksheet, whether or not it yielded a table, with the rectangle the
  read covered and counts of the cells in it that a `Table` cannot represent.
  `uncachedFormulaCells` counts the cells holding a formula the workbook carries
  no calculated value for: Excel writes a formula and its last result together,
  but a file written by a generator carries the formula alone, and every reader
  then sees those cells as empty because empty is all the file says.
  `errorCells` counts the cells holding an error value, typed in or left by a
  formula whose last calculation failed: those are stored as `t="e"` and the
  spreadsheet engine reports the internal code Excel numbers each error by, so a
  `#REF!` arrives as `23` and a `#DIV/0!` as `7`, and a column of amounts still
  infers a numeric type. Both surfaces adapt their input and call one operation,
  so neither can answer differently about the same workbook. Both counts are
  read from this package's own document model, because the spreadsheet engine
  drops an uncalculated cell while parsing and flattens an error into that code,
  and they are taken in one walk over the rectangle the table reader reported
  rather than over a second header resolution, so they can only describe the
  read the caller got. Whether a cached value is present is decided by the
  element being there rather than by what it holds, so a formula that evaluated
  to an empty string or to zero counts as calculated; that is the definition the
  values-only conversion already used, now shared between them. Reading,
  consolidating, merging, and splitting are unchanged and still treat such a
  cell as empty or as its code.

  A worksheet part is parsed the first time something asks for it, so a
  malformed row or cell reference surfaced from an ordinary property access as a
  bare parser error while the same damage in the package surfaced as
  `XLSX_READ_FAILED`. Both now go through one translation: a failed read reports
  `XLSX_READ_FAILED` naming the workbook and, where the failure belongs to a
  worksheet, that worksheet, with the parser's own complaint as the error's
  `cause`.

  Dates read out of a workbook no longer depend on the reader's time zone. This
  is a correctness fix, and it changes what every non-UTC caller gets.

  Excel holds a date as a count of days from the workbook's epoch, wearing a
  date number format, and attaches no zone. Three readers turned that into a
  JavaScript `Date` and then took its ISO 8601 face, and a `Date` has a local
  face as well as a UTC one:

  - The document model assembled the components in local time, so serial 45292
    read as `2024-01-01T00:00:00Z` in UTC, `2023-12-31T20:00:00Z` in UTC+4 and
    `2023-12-31T18:30:00Z` in UTC+5:30. East of UTC the calendar day moved to
    the one before the one in the cell.
  - A `t="d"` cell writes ISO text with no zone on it, and `new Date(text)` read
    that as a local moment: `2024-01-01T18:00:00` became 14:00Z in UTC+4 and the
    next day in UTC-7.
  - The worksheet readers asked the spreadsheet engine for a `Date` and rendered
    it, which was correct only because that engine happens to compose the
    instant from a UTC epoch. Nothing said so, and a change underneath would
    have moved every date silently.

  All three now work from the workbook's own numbers: the readers decode the
  serial and the workbook's `date1904` flag into calendar components with the
  engine's `parse_date_code` and write the components out, and the model
  composes its `Date` with `Date.UTC`. The text is a function of the workbook
  alone, and it is always the full timestamp, `2024-01-01T00:00:00.000Z`,
  whether or not the cell carries a time, so a date column reads one way down
  its length and a value the workbook holds as a date stays distinguishable from
  text somebody typed. `@consultchimps/db` accepts that spelling in a `date`
  column.

  A `t="d"` cell, which writes its date as ISO 8601 text, is now read into
  components, judged as one value, and converted by arithmetic. It used to reach
  `Date.UTC`, which carries rules of its own: a year from 0 to 99 is remapped
  into the twentieth century, so `0099-01-01` read back as 1999, and an
  out-of-range field is normalised rather than refused, so `2024-13-01` read
  back as January 2025. A month is now 1 to 12, a day has to exist in that month
  of that year, hour is 0 to 23 and minute and second 0 to 59, and a written
  offset is 0 to 23 hours and 0 to 59 minutes. Text that names no moment is
  carried as the text the cell holds, unconverted, the way every other
  unconvertible cell is; nothing is guessed and nothing is normalised.

  Two edges of that reading are held to the same rule. A written offset moves
  the moment it names, and it can move it out of the years that can be written:
  `9999-12-31T23:30:00-01:00` is a good timestamp whose UTC face is the year
  10000, which ISO 8601 spells only in an expanded form that neither a `date`
  column nor an output filename reads. The range rule now judges the adjusted
  moment as well as the components as written, and a value that leaves the range
  is carried as its text. The range is 0000 to 9999, the years four digits
  spell, which is exactly what `isIsoDateText` accepts; both ends are named in a
  test on either side of that seam.

  And a serial's time of day is rounded as one value before it is split, not
  field by field afterwards. A serial naming 23:59:59.9996 rounds to a whole
  day, which carries into midnight of the next day; rounded after the second had
  already been decided, the carry had nowhere to go, the fraction was clamped to
  999 milliseconds, and the value read as the second before, on the day before.

  Turning a serial into a moment is one route now, and it judges what it makes.
  There were two: the reader asked the spreadsheet engine, and the document
  model did its own arithmetic. They disagreed about which day serial 1 is,
  because the model counted from an epoch that only holds for serials past the
  day the 1900 system invents, so every date before 1 March 1900 read one day
  early through a split's group keys. Neither judged the result: a
  date-formatted cell holding an untrusted 1e100 became a moment with no
  components at all, spelled `0NaN-NaN-NaNTNaN:NaN:NaN.NaNZ`, so distinct
  serials were keyed alike and a split gathered them into one output workbook.
  Both now decode through `serialCalendarParts`, which rounds the time of day as
  one whole number of milliseconds, derives the calendar day by the module's own
  arithmetic, and refuses anything that names no moment; a serial that names
  none is carried as the number it is, the same decision text that names no
  moment gets. The unvalidated conversion is gone rather than deprecated, so no
  caller can spell a moment nothing judged.

  One consequence worth naming: serial 60 is carried as a number rather than
  written as a date. It is the day the 1900 system invents, 29 February 1900,
  which the Gregorian calendar does not have, no `Date` can hold, and a `date`
  column in `@consultchimps/db` refuses. Writing it made the reader and the
  split key report two different days for one cell.

  A worksheet says a cell is a date in two ways, and both are read now. Excel
  writes a count of days wearing a date number format; the format also lets a
  cell declare `t="d"` and write ISO 8601 text, which other generators use and
  Excel opens. The worksheet readers saw only the first, because reading with
  the engine's dates off, which is what keeps a serial a serial, turns a
  declared date into a plain number with no field left saying what it was:
  measured on the pinned engine, `<c t="d"><v>2024-01-01</v></c>` arrives
  as 45292. So such a cell read as an integer, and an import inferred an integer
  column for it, while the document model read the same cell correctly. The
  readers now take dates from the model, which is the reader that sees the
  declared type, owns the style table, and goes through the one calendar route;
  the engine keeps the used range, the header row, and every cell that is not a
  date. Reading the type from the engine instead was measured and rejected: with
  its dates on it keeps the declaration but hands back its own parse of the
  text, which remaps a year of 0099 to 1999 and normalises a month of 13 into
  January of the next year.

  A declared date cell holding nothing, or nothing but spaces, is blank. The
  branch that reads one skipped the rule its neighbour applies to every other
  cell and handed back the empty text, which is a value: an empty declared date
  above a worksheet's real header counted as content, so the header row was
  found a row early, the columns became invented names and the real header
  imported as data. The rule now lives in the one switch that reads a cell,
  beside the identical rule for a numeric cell two lines below it. A number that
  is not finite is blank too, which is what the engine makes of text it cannot
  read.

  Measuring that scenario turned up something worse and it is fixed here as
  well. The engine cannot parse a worksheet containing a declared date with no
  text at all: it lists the sheet, produces nothing for it, and raises nothing.
  The reader skipped such a worksheet, so it vanished from every list it feeds -
  the tables, the worksheets an import offers - and nothing said so. That is now
  a refusal with the stable read code, naming the workbook and the worksheet.
  The document model reads those worksheets, so this is a limit of the reader
  rather than of the file, and it goes when the reader takes its cells from the
  model.

  A declared date whose text names no moment is carried as that text, as the
  model already did, and so is one written finer than a millisecond:
  `18:00:00.1234` and `18:00:00.1239` were both shortened to `.123`, which
  changed the value on the way in and let a split gather rows that were not
  together. Digits past the third that are zeros lose nothing and are still
  accepted. The reader now states its whole contract beside the grammar - every
  shape the text can arrive in and what each one answers - and a test pins the
  enumeration line for line; the serial path states and pins its own. Writing it
  down settled two shapes that had no stated answer: a lower case `t` separator
  is read, as a lower case `z` already was, and an offset written without
  minutes is carried as text, which is what the profile this format names
  allows. Displayed text, which `readWorksheetRecords` reports, still comes from
  the engine for a cell that wears a date format, because that is the text the
  worksheet shows; for a declared date it comes from the model, because the
  engine's text for one is the serial it made of it.

  One spelling of a date, `calendarIsoText`, is now shared by the worksheet
  readers, the document model and the split's group keys, so a value read from a
  cell and the key that names the output workbook it lands in cannot describe
  one cell two ways.

  The blast radius is every consumer that reads a cell through these readers:
  `readWorkbookTables`, `readWorkbookWorksheets`, the Excel Table and
  named-range readers, worksheet records, consolidation, and the split, whose
  group keys become output workbook filenames. In UTC nothing changes; in every
  other zone these now report the day the cell names. Tests read the same
  workbook with the host in five zones, from UTC-7 to UTC+14, and require one
  answer, in both of Excel's date systems.

  A row or cell that carries no `r` attribute keeps its place through an edit.
  Both are optional in the format: such a row or cell sits where document order
  puts it. The model infers the position and writes it back, because an implicit
  position is only true of the document it was read from, but it wrote it with
  the helper that updates an attribute already there, which returns the tag
  unchanged when there is none. The number therefore existed in memory and never
  in the output: filtering an Excel Table compacted the rows above such a row
  while it stayed implicit, so it serialised directly under the shortened table
  and Excel read it as a different row. The write back now uses a helper whose
  contract is the operation it needed - the tag carries this value afterwards,
  whether or not it had the attribute - and the tests assert the serialised tag
  rather than the parsed model.

  `uniqueHeaders` in `@consultchimps/tabular` now decides uniqueness against the
  whole header row: every original spelling is reserved before any suffix is
  generated, so a row such as `A, A, A_2` keeps three distinct columns instead
  of writing two of them under one name and silently dropping a column's values.
  A run of repeated headers is named in one pass rather than re-probing every
  suffix already handed out. Every reader that names columns (delimited text,
  worksheets, Excel Tables, named ranges, and the workbook description) inherits
  both, which is why the package is versioned here.

### Patch Changes

- Updated dependencies [04dd6de]
  - @consultchimps/tabular@0.4.1

## 0.17.0

### Minor Changes

- 1eec990: Bring the workbook unprotect operation up to the package standard.
  The XML edit now runs on the package layer through a new
  `WorkbookPackage.removeEmptyElements` seam instead of a regex inside the
  operation, unprotect gains a declared column in the conformance contract
  (every tracked structure is `preserve`, held up by a new corpus), and
  unprotect now refuses an output name whose extension contradicts the
  workbook's declared type with the new stable error
  `XLSX_UNPROTECT_PACKAGE_TYPE_MISMATCH`, before writing anything.

## 0.16.0

### Minor Changes

- 9444c7f: Add byte-preserving Excel worksheet and workbook-structure
  unprotection for `.xlsx` and `.xlsm` files.

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
