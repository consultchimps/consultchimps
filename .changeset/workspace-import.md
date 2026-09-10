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

`suggestTableName` offers a cleaned name at the full identifier limit. It used
to leave room for the `table_` prefix in every suggestion, including the ones
nothing would ever prefix, so a name between 195 and 200 characters came back
shortened for a prefix it was not going to get, and two files differing only in
those last characters were handed the same name to resolve by hand. The room is
now taken in the branch that adds the prefix, from the one name that needs it,
and the result fits by construction either way.

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

Dates read out of a workbook no longer depend on the reader's time zone. This is
a correctness fix, and it changes what every non-UTC caller gets.

Excel holds a date as a count of days from the workbook's epoch, wearing a date
number format, and attaches no zone. Three readers turned that into a JavaScript
`Date` and then took its ISO 8601 face, and a `Date` has a local face as well as
a UTC one:

- The document model assembled the components in local time, so serial 45292
  read as `2024-01-01T00:00:00Z` in UTC, `2023-12-31T20:00:00Z` in UTC+4 and
  `2023-12-31T18:30:00Z` in UTC+5:30. East of UTC the calendar day moved to the
  one before the one in the cell.
- A `t="d"` cell writes ISO text with no zone on it, and `new Date(text)` read
  that as a local moment: `2024-01-01T18:00:00` became 14:00Z in UTC+4 and the
  next day in UTC-7.
- The worksheet readers asked the spreadsheet engine for a `Date` and rendered
  it, which was correct only because that engine happens to compose the instant
  from a UTC epoch. Nothing said so, and a change underneath would have moved
  every date silently.

All three now work from the workbook's own numbers: the readers decode the
serial and the workbook's `date1904` flag into calendar components with the
engine's `parse_date_code` and write the components out, and the model composes
its `Date` with `Date.UTC`. The text is a function of the workbook alone, and it
is always the full timestamp, `2024-01-01T00:00:00.000Z`, whether or not the
cell carries a time, so a date column reads one way down its length and a value
the workbook holds as a date stays distinguishable from text somebody typed.
`@consultchimps/db` accepts that spelling in a `date` column.

A `t="d"` cell, which writes its date as ISO 8601 text, is now read into
components, judged as one value, and converted by arithmetic. It used to reach
`Date.UTC`, which carries rules of its own: a year from 0 to 99 is remapped into
the twentieth century, so `0099-01-01` read back as 1999, and an out-of-range
field is normalised rather than refused, so `2024-13-01` read back as
January 2025. A month is now 1 to 12, a day has to exist in that month of that
year, hour is 0 to 23 and minute and second 0 to 59, and a written offset is 0
to 23 hours and 0 to 59 minutes. Text that names no moment is carried as the
text the cell holds, unconverted, the way every other unconvertible cell is;
nothing is guessed and nothing is normalised.

Two edges of that reading are held to the same rule. A written offset moves the
moment it names, and it can move it out of the years that can be written:
`9999-12-31T23:30:00-01:00` is a good timestamp whose UTC face is the year
10000, which ISO 8601 spells only in an expanded form that neither a `date`
column nor an output filename reads. The range rule now judges the adjusted
moment as well as the components as written, and a value that leaves the range
is carried as its text. The range is 0000 to 9999, the years four digits spell,
which is exactly what `isIsoDateText` accepts; both ends are named in a test on
either side of that seam.

And a serial's time of day is rounded as one value before it is split, not field
by field afterwards. A serial naming 23:59:59.9996 rounds to a whole day, which
carries into midnight of the next day; rounded after the second had already been
decided, the carry had nowhere to go, the fraction was clamped to 999
milliseconds, and the value read as the second before, on the day before.

Turning a serial into a moment is one route now, and it judges what it makes.
There were two: the reader asked the spreadsheet engine, and the document model
did its own arithmetic. They disagreed about which day serial 1 is, because the
model counted from an epoch that only holds for serials past the day the 1900
system invents, so every date before 1 March 1900 read one day early through a
split's group keys. Neither judged the result: a date-formatted cell holding an
untrusted 1e100 became a moment with no components at all, spelled
`0NaN-NaN-NaNTNaN:NaN:NaN.NaNZ`, so distinct serials were keyed alike and a
split gathered them into one output workbook. Both now decode through
`serialCalendarParts`, which rounds the time of day as one whole number of
milliseconds, derives the calendar day by the module's own arithmetic, and
refuses anything that names no moment; a serial that names none is carried as
the number it is, the same decision text that names no moment gets. The
unvalidated conversion is gone rather than deprecated, so no caller can spell a
moment nothing judged.

One consequence worth naming: serial 60 is carried as a number rather than
written as a date. It is the day the 1900 system invents, 29 February 1900,
which the Gregorian calendar does not have, no `Date` can hold, and a `date`
column in `@consultchimps/db` refuses. Writing it made the reader and the split
key report two different days for one cell.

A worksheet says a cell is a date in two ways, and both are read now. Excel
writes a count of days wearing a date number format; the format also lets a cell
declare `t="d"` and write ISO 8601 text, which other generators use and Excel
opens. The worksheet readers saw only the first, because reading with the
engine's dates off, which is what keeps a serial a serial, turns a declared date
into a plain number with no field left saying what it was: measured on the
pinned engine, `<c t="d"><v>2024-01-01</v></c>` arrives as 45292. So such a cell
read as an integer, and an import inferred an integer column for it, while the
document model read the same cell correctly. The readers now take dates from the
model, which is the reader that sees the declared type, owns the style table,
and goes through the one calendar route; the engine keeps the used range, the
header row, and every cell that is not a date. Reading the type from the engine
instead was measured and rejected: with its dates on it keeps the declaration
but hands back its own parse of the text, which remaps a year of 0099 to 1999
and normalises a month of 13 into January of the next year.

A declared date cell holding nothing, or nothing but spaces, is blank. The
branch that reads one skipped the rule its neighbour applies to every other cell
and handed back the empty text, which is a value: an empty declared date above a
worksheet's real header counted as content, so the header row was found a row
early, the columns became invented names and the real header imported as data.
The rule now lives in the one switch that reads a cell, beside the identical
rule for a numeric cell two lines below it. A number that is not finite is blank
too, which is what the engine makes of text it cannot read.

Measuring that scenario turned up something worse and it is fixed here as well.
The engine cannot parse a worksheet containing a declared date with no text at
all: it lists the sheet, produces nothing for it, and raises nothing. The reader
skipped such a worksheet, so it vanished from every list it feeds - the tables,
the worksheets an import offers - and nothing said so. That is now a refusal
with the stable read code, naming the workbook and the worksheet. The document
model reads those worksheets, so this is a limit of the reader rather than of
the file, and it goes when the reader takes its cells from the model.

A declared date whose text names no moment is carried as that text, as the model
already did, and so is one written finer than a millisecond: `18:00:00.1234` and
`18:00:00.1239` were both shortened to `.123`, which changed the value on the
way in and let a split gather rows that were not together. Digits past the third
that are zeros lose nothing and are still accepted. The reader now states its
whole contract beside the grammar - every shape the text can arrive in and what
each one answers - and a test pins the enumeration line for line; the serial
path states and pins its own. Writing it down settled two shapes that had no
stated answer: a lower case `t` separator is read, as a lower case `z` already
was, and an offset written without minutes is carried as text, which is what the
profile this format names allows. Displayed text, which `readWorksheetRecords`
reports, still comes from the engine for a cell that wears a date format,
because that is the text the worksheet shows; for a declared date it comes from
the model, because the engine's text for one is the serial it made of it.

One spelling of a date, `calendarIsoText`, is now shared by the worksheet
readers, the document model and the split's group keys, so a value read from a
cell and the key that names the output workbook it lands in cannot describe one
cell two ways.

The blast radius is every consumer that reads a cell through these readers:
`readWorkbookTables`, `readWorkbookWorksheets`, the Excel Table and named-range
readers, worksheet records, consolidation, and the split, whose group keys
become output workbook filenames. In UTC nothing changes; in every other zone
these now report the day the cell names. Tests read the same workbook with the
host in five zones, from UTC-7 to UTC+14, and require one answer, in both of
Excel's date systems.

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
