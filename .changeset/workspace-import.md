---
"@consultchimps/db": minor
"@consultchimps/xlsx": minor
"@consultchimps/tabular": patch
---

Import a table of values into a local database.

`@consultchimps/db` gains the import half of the workspace, as runtime-neutral
public API: `inferColumnTypes` decides each column's type from its own values,
`importTable` and `importTables` create the tables and fill them with generated
Record IDs, `importedTableSchema` shows what an import would create without
creating it, and `suggestTableName` and `suggestRecordIdPrefix` derive the
starting points a person edits before committing. `parseCsvTable` reads
comma-separated text (RFC 4180 quoting, CRLF, LF, and CR line endings, a leading
byte-order mark) as a `Table` of text values, so the type of a column is decided
in one place rather than guessed by a spreadsheet engine while it parses.

Inference is conservative and never guesses: a column takes `boolean`, `date`,
`integer`, or `real` only when every value that is not blank fits it, and stays
`text` otherwise. Both numeric rules run one round-trip test, so a value is read
as a number only when it reads back as the number it wrote: a padded reference
code such as `007`, an exponent form, a whole number past the range stored
exactly, and a decimal carrying more digits than a number can hold all stay
text. A timestamp is judged as one value, so hour 24 is accepted only as the end
of a day, second 60 only as a leap second at 23:59, and an offset only within
its real range. An import of several tables is one unit: names, Record ID
prefixes, and column names are checked first, then every table is created and
filled inside a single transaction, so a failure part way through leaves the
database exactly as it was.

A `date` column now holds ISO 8601 text and nothing else. `sqlValueFromCell`,
the single point every write goes through, judges a date with `isIsoDateText`,
the same rule the import's inference uses to decide a column is a date, and
refuses anything else with `DB_INVALID_DATE`; reading back a stored value that
is not one reports damage. Surrounding spaces are removed before the value is
judged and stored, and a value of nothing but spaces is stored as null, so a
column that claims a spelling holds it and an empty cell is findable as empty.
This is a behaviour change: a date column that previously accepted any text now
accepts only a date.

A source whose columns repeat a name exactly is refused with
`DB_IMPORT_DUPLICATE_SOURCE_COLUMN`. A `Table` row is an object keyed by column
name, so the second `Amount` in `["Amount", "Amount"]` is a column no row can
answer: the remap read the one property twice and stored the same value under
two names, manufacturing a duplicate the source never had. Names differing only
in case, or in what the identifier fold ignores, are two properties carrying two
values and are still numbered and reported, not refused. Every reader here makes
its headers unique, so this reaches only a `Table` a caller built.

`importedTableSchema` now refuses what the import refuses. The column, table
name and Record ID checks ran beside the shared plan rather than inside it, so a
source with no columns, or one whose only column is the Record ID, previewed as
a schema with no columns while the import raised `DB_IMPORT_NO_COLUMNS` and
created nothing, and a preview would describe a table under a name the engine
will not take. They live in the plan now; only the checks that need the
workspace, such as a name it already holds, still belong to the import.

A column name is derived in the same step that numbers repeated ones, so it is
inside the identifier limit by construction: a header too long to be a name is
shortened at a whole-character boundary with room left for the number a
duplicate may need, and `ImportedTable.renamedColumns` reports every header that
ended up under a different name. Shortening after numbering, in a different
layer, refused a whole import when two copies of a 200-character header became a
name of 202.

Also new: `Database.countRecords` counts a table's records in the engine rather
than by reading its rows, `assertRecordIdConfig` and `MAX_RECORD_ID_PADDING`
make the Record ID rules callable, `MAX_IDENTIFIER_LENGTH` and
`truncateIdentifier` state the identifier limit and shorten a name to fit it
without ever cutting a character in half, and `@consultchimps/db/schema`
publishes the schema model on its own, with no database engine behind it, so a
browser page can fold an identifier without downloading WebAssembly to do it.
`assertSafeIdentifier` now also refuses a name that is not well-formed text: an
incomplete character passed every other check and was stored as U+FFFD, so the
name in the database was not the name that was asked for.

`@consultchimps/xlsx` gains `readWorkbookTablesBytes`, the byte twin of
`readWorkbookTables`: every visible worksheet that holds data as a `Table`, with
the same `sheets`, `headerRow`, and `includeHiddenSheets` options. The byte
surface had readers for Excel Tables and named ranges but none for the
worksheets themselves.

`readWorkbookWorksheets` and `readWorkbookWorksheetsBytes` join it: every
selected worksheet, whether or not it yielded a table, with the rectangle the
read covered and counts of the cells in it that a `Table` cannot represent.
`uncachedFormulaCells` counts the cells holding a formula the workbook carries
no calculated value for: Excel writes a formula and its last result together,
but a file written by a generator carries the formula alone, and every reader
then sees those cells as empty because empty is all the file says. `errorCells`
counts the cells holding an error value, typed in or left by a formula whose
last calculation failed: those are stored as `t="e"` and the spreadsheet engine
reports the internal code Excel numbers each error by, so a `#REF!` arrives as
`23` and a `#DIV/0!` as `7`, and a column of amounts still infers a numeric
type. Both surfaces adapt their input and call one operation, so neither can
answer differently about the same workbook. Both counts are read from this
package's own document model, because the spreadsheet engine drops an
uncalculated cell while parsing and flattens an error into that code, and they
are taken in one walk over the rectangle the table reader reported rather than
over a second header resolution, so they can only describe the read the caller
got. Whether a cached value is present is decided by the element being there
rather than by what it holds, so a formula that evaluated to an empty string or
to zero counts as calculated; that is the definition the values-only conversion
already used, now shared between them. Reading, consolidating, merging, and
splitting are unchanged and still treat such a cell as empty or as its code.

A worksheet part is parsed the first time something asks for it, so a malformed
row or cell reference surfaced from an ordinary property access as a bare parser
error while the same damage in the package surfaced as `XLSX_READ_FAILED`. Both
now go through one translation: a failed read reports `XLSX_READ_FAILED` naming
the workbook and, where the failure belongs to a worksheet, that worksheet, with
the parser's own complaint as the error's `cause`.

A row or cell that carries no `r` attribute keeps its place through an edit.
Both are optional in the format: such a row or cell sits where document order
puts it. The model infers the position and writes it back, because an implicit
position is only true of the document it was read from, but it wrote it with the
helper that updates an attribute already there, which returns the tag unchanged
when there is none. The number therefore existed in memory and never in the
output: filtering an Excel Table compacted the rows above such a row while it
stayed implicit, so it serialised directly under the shortened table and Excel
read it as a different row. The write back now uses a helper whose contract is
the operation it needed - the tag carries this value afterwards, whether or not
it had the attribute - and the tests assert the serialised tag rather than the
parsed model.

`uniqueHeaders` in `@consultchimps/tabular` now decides uniqueness against the
whole header row: every original spelling is reserved before any suffix is
generated, so a row such as `A, A, A_2` keeps three distinct columns instead of
writing two of them under one name and silently dropping a column's values. A
run of repeated headers is named in one pass rather than re-probing every suffix
already handed out. Every reader that names columns (delimited text, worksheets,
Excel Tables, named ranges, and the workbook description) inherits both, which
is why the package is versioned here.
