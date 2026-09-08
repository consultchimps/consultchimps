---
"@consultchimps/db": minor
"@consultchimps/xlsx": minor
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

Also new: `Database.countRecords` counts a table's records in the engine rather
than by reading its rows, `assertRecordIdConfig` and `MAX_RECORD_ID_PADDING`
make the Record ID rules callable, and `@consultchimps/db/schema` publishes the
schema model on its own, with no database engine behind it, so a browser page
can fold an identifier without downloading WebAssembly to do it.

`@consultchimps/xlsx` gains `readWorkbookTablesBytes`, the byte twin of
`readWorkbookTables`: every visible worksheet that holds data as a `Table`, with
the same `sheets`, `headerRow`, and `includeHiddenSheets` options. The byte
surface had readers for Excel Tables and named ranges but none for the
worksheets themselves.
