# Power BI table export

Status: Proposed. The initial decisions followed a reading spike. The
clarifications below address review findings before acceptance.

Consultants receive Power BI files (`.pbix`) whose data they need in Excel, and
the only supported way to get it out is to open the file in Power BI Desktop on
Windows and query the running engine. This ADR adds the toolkit's first Power BI
operation: export each exportable table of a `.pbix` file's model to a workbook,
as one or more worksheets per table, entirely in the browser, with the file
never leaving the machine.

A `.pbix` is a zip. The loaded rows live in its `DataModel` part: an Analysis
Services backup, compressed with Microsoft's XPress9 algorithm, holding a SQLite
catalog and one VertiPaq column store per column. An earlier spike rejected
reading these rows because no JavaScript XPress9 decompressor exists. That was
the wrong conclusion drawn from a true fact. Microsoft's reference decompressor
is MIT-licensed C, and C compiles to WebAssembly the same way the sql.js engine
the toolkit already ships does. A second spike (2026-09-10, ten public Microsoft
sample files) proved the whole pipeline: nine of nine `.pbix` files, sixty-two
tables and 808,559 rows decoded cell for cell identically to an independent
reader (pbixray), in a Web Worker in headless Chromium with no cross-origin
isolation headers, then written to byte-deterministic workbooks that a
third-party reader round-tripped with no differences. The tenth file, a `.pbit`
template, carries no model and is the refusal case.

Query definitions (the Power Query M code) are a different reading and are not
in scope: modern `.pbix` files no longer carry them as a separate part, so that
feature would only work on files from before late 2020 or on templates. Rows are
what the consultant asked for, and rows work on current files.

## Decision 1: a new package, `@consultchimps/pbi`

Options considered: a new package; folding the reader into `@consultchimps/xlsx`
as another input format.

**Decision: a new package `@consultchimps/pbi`.** It carries a WebAssembly
artifact and a vendored third-party C source tree, and neither belongs in a
package whose consumers expect pure TypeScript. It depends on
`@consultchimps/db` for the sql.js boundary (the catalog inside the model is
SQLite) and on `@consultchimps/xlsx` for the workbook writer. No new npm
dependency enters the graph: the runtime stack is sql.js and jszip at the exact
versions the repository already pins.

## Decision 2: only the decompressor is WebAssembly, and its binary is committed

The XPress9 decoder is three C files from Microsoft (under MIT, obtained from
the `xpress9-python` repository at a recorded commit) plus a short shim, built
single-threaded with Emscripten to about 56 KB (20 KB gzipped). It needs no
threads and no `SharedArrayBuffer`, so GitHub Pages can serve it without the
cross-origin isolation headers it cannot set. Everything above the decompressor,
the backup container, the catalog query, the Huffman string dictionaries and the
VertiPaq column decoding, is plain TypeScript. The spike's Huffman kernel port
showed that none of it needs native code.

Options considered for the build: commit the built binary; make Emscripten a
prerequisite of `pnpm build`; track the C source as a git submodule.

**Decision: vendor the C source with its upstream commit hash recorded, commit
the built `.wasm`, and have CI rebuild it and assert identical bytes.**
Contributors do not need a 700 MB toolchain to build the repository, the
reproducibility check keeps the committed binary honest, and a vendored copy is
safer than a submodule pointing at one person's repository. The binary is served
from our own origin as a static asset, never a CDN, exactly as the sql.js binary
is. The Microsoft copyright and MIT text are reproduced in
`THIRD-PARTY-LICENSES.md` and beside the vendored source, and the ported Huffman
kernel attributes its origin in its file header.

Before committing the first binary, build item 2 pins the Emscripten toolchain
image by immutable digest, including its compiler, linker, and optimizer. A
committed build script records the complete compiler and linker flags, fixes the
build path and timestamp inputs, and is the single command used locally and in
CI. No floating toolchain tags or runner-provided compiler are used. The check
builds twice in clean copies with that pinned environment and compares both
outputs with the committed binary. Toolchain upgrades update the pin and rebuilt
binary together in a reviewed PR.

## Decision 3: a named tool in a new "Power BI" category

Options considered: a named registry operation with its own category; a `.pbix`
input option on the data workspace import; both.

**Decision: a named operation, `pbi.export-tables`, in a new "Power BI" tool
category, with its own guide page and browser tool page.** The output is a
workbook, which argues for treating `.pbix` as just another input, but the
refusal surface is large and specific to Power BI (templates, live connections,
DirectQuery, encrypted models, unsupported encodings), and a tool that refuses
needs a page that can explain why. The category is the fourth value of
`ToolCategory` in the registry, a schema change that touches the README
operations table and the registry drift checks (ADR 0001 is otherwise
unchanged). Adding `.pbix` to workspace import can follow later as a separate
decision.

## Decision 4: refusal and partial export

**Decision: the operation has three model-policy refusals, each with its own
stable code: the file holds no model, the model is encrypted, or the model has
no exportable table. A readable model with at least one exportable table
produces a workbook when it passes the capacity limits below, and every excluded
table or column is reported in a manifest.** Before policy applies, input that
cannot be read at all is refused with a validation code rather than a leaked
parser error: a file that is not a zip or lacks the parts a `.pbix` must have
(`PBI_INVALID_CONTAINER`), and a model part whose compressed stream, backup
container, or catalog is truncated or corrupt (`PBI_MODEL_UNREADABLE`). A valid
Power BI container without a `DataModel` part is refused with `PBI_NO_MODEL`. A
missing `DataModel` is not an invalid container. The message says the file has
no embedded model to export and asks for a `.pbix` saved with imported data.
Templates, live connections, and DirectQuery-only files are possible causes, not
diagnoses inferred from an extension or a missing part. They share this code
until fixtures establish reliable distinctions. The presence of a `Connections`
part is not a live-connection signal: a current Microsoft sample carries one
alongside a full model, and the spike hit that false positive. An encrypted or
password-protected model is refused (`PBI_MODEL_ENCRYPTED`); the reader does not
attempt it.

When a file decodes partially, for example one column uses an encoding the
reader has not seen, the operation still produces the workbook and lists each
skipped table or column in the manifest with a stable code. A table remains
exportable when at least one column can be decoded and represented. Retained
columns keep their row positions and nulls; dropping a column never shifts or
drops rows. A column whose complete values or row alignment cannot be recovered
is skipped as a whole. If no columns remain, the table is excluded. A table with
zero rows and at least one output column produces a header-only worksheet.
Successful partial exports also populate `OperationResult.warnings` from the
manifest under Decision 9, so the normal result summary reports the exclusions.

A workbook is only produced when at least one table is exportable. When the
model holds tables but every one of them is excluded, whether hidden under the
default policy, over the column limit, or without any usable output columns, the
operation fails with a stable code (`PBI_NO_EXPORTABLE_TABLES`) whose message
lists each table with the reason it was excluded, and names the include-hidden
option when that is the cure. An empty workbook is never a success.

Capacity is a separate operational gate, with `PBI_EXPORT_LIMIT_EXCEEDED` for an
export that cannot fit the configured limits. The shared byte engine accepts
positive, safe-integer byte limits for input bytes, decoded bytes, output bytes,
and estimated peak working memory. Omitted limits use the shipped defaults.
Validate supplied limits before reading input or allocating export buffers.
Reject a wrong type, zero, negative value, fraction, unsafe integer, `NaN`, or
infinity with `PBI_INVALID_OPTIONS`. Its details list each invalid option name
and the requirement for a positive safe integer, in the limit order above; the
message tells the caller to omit the option or supply a valid byte count. These
option errors produce no artifacts or destination writes and are distinct from
exceeding a valid limit. Browser defaults are fixed, measured values shipped
with the implementation, not guesses based on available device memory. Build
items 1 and 8 must establish the limits and their estimator fixtures before item
9 can claim browser support. Larger limits require an explicit caller option and
are outside the default browser envelope.

Check the input size before reading a browser `File` into an array. Use
validated container sizes, model row counts, and dictionary/value widths to
conservatively bound decompression, table cells, workbook structures,
serialization buffers, the manifest, the combined-download archive, and
worker-to-host copies before allocating them. Reserve the worst case, including
uncompressed output; compressed file size alone is never the estimate. If a
supported format path cannot supply a safe upper bound, refuse with the same
capacity code instead of attempting an unbounded export. Enforce cumulative byte
and allocation budgets during decompression and writing as well, since file
metadata is untrusted. Check each growth before allocating; do not depend on
catching an out-of-memory exception. The error identifies the stage, configured
limit, and either the required bound or why it cannot be estimated. It suggests
a smaller model. Capacity refusal returns no byte artifacts and occurs before a
file adapter writes destinations.

## Decision 5: tables larger than a worksheet

Excel holds 1,048,576 rows per worksheet and Power BI models routinely exceed
that. Options considered: refuse the table and export the rest; split the table
across numbered worksheets; export the first million rows with a manifest entry.

**Decision: split the table across numbered worksheets** (`Sales`, `Sales_2`,
`Sales_3`), with the split recorded in the manifest. The decision-maker chose
this over refusing the table, so that no rows are ever withheld; the cost,
accepted knowingly, is that lookups and pivots across the parts are the user's
job. Every part carries the header row, so a part holds at most 1,048,575 data
rows (the worksheet limit less the header), and the last part is never asked to
hold a row the worksheet cannot address. Column count above 16,384 refuses the
table (such a model is pathological). Ordinary cell text is limited to 32,767
UTF-16 code units and 253 line breaks. Count a CRLF pair as one break and a lone
CR or LF as one break; retain the original line-ending characters. Truncate to
the longest prefix satisfying both limits without splitting a surrogate pair or
a CRLF pair. In particular, stop before the 254th line break even when the
code-unit limit has not been reached. Count one `PBI_TEXT_TRUNCATED` per changed
cell, including a cell exceeding both limits, not per removed code unit or line
break. Binary encodings follow Decision 6 instead and are never truncated.
Worksheet names follow Excel's rules (31 characters, forbidden characters
replaced, case-insensitive uniqueness with numeric suffixes).

Text serialization must preserve the resulting string, including controls and
literal SpreadsheetML escape sequences. Measure and truncate the logical text
before escaping; encoded XML length does not consume the cell's text allowance.
At the workbook serialization boundary, apply
[SpreadsheetML string escaping](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/d34ae755-c53f-4a44-a363-c6dd3ee018a4)
with lowercase `x` and four uppercase hex digits. Encode XML-disallowed control
characters and carriage returns as `_xHHHH_`. Protect an original literal
`_xHHHH_`, with either case of hex digits, by encoding its leading underscore as
`_x005F_`; do not re-escape sequences generated by this pass. Thus literal
`_x000A_` serializes as `_x005F_x000A_`, while an actual line feed remains a
line feed in element text. Apply ordinary XML entity escaping as well, and
preserve leading and trailing whitespace. The workbook writer owns this
encoding; callers pass logical strings and must not pre-escape them.

Escaping is lossless and adds no manifest conversion reason. All text cells,
including headers and Decision 6 text representations, must read back as their
intended logical strings through an independent workbook reader. Build item 8
must prove this before browser exposure, with fixtures for NUL and other
controls, CR/LF/CRLF, literal escapes including `_x005F_`, XML metacharacters,
edge whitespace, supplementary Unicode, and truncation adjacent to these
characters. Pin expected XML bytes as well as read-back values so an escaping
change cannot silently change the deterministic output contract.

Worksheet allocation is deterministic. Tables are processed in ascending numeric
catalog `Table.ID` order, never in decode-completion order, and each table
claims all of its worksheet names before the next table is considered: its
sanitized name first, then its numbered parts. A name already claimed,
case-insensitively, whether by an earlier table's name or one of its parts,
takes the next free numeric suffix. So a real table named `Sales_2` that follows
a split `Sales` becomes `Sales_2_2`, and the same file always yields the same
names, tab order, and bytes. Within a table, columns follow ascending
`ColumnStorage.StoragePosition`, then numeric `Column.ID` as a tie-breaker.
These are the table and column catalog orders used throughout this ADR.

Within each table, concatenate partitions by ascending
`PartitionStorage.StoragePosition`, then numeric `PartitionStorage.ID` as a
tie-breaker. This uses the stored partition position exposed by the
[reference catalog reader](https://github.com/Hugoberry/pbixray/blob/c5da940b45d2806c8f8d3e109f0d1909117a7db0/pbixray/meta/sqlite_source.py#L281).
Within a partition, concatenate segments by their zero-based position in its
stored segment-descriptor sequence, then retain the decoded row order inside
each segment. All retained columns must use that same partition and segment
sequence and aligned row counts. Apply `PBI_COLUMN_ROW_ALIGNMENT_UNRECOVERABLE`
when a column cannot be placed on that sequence. Neither catalog query return
order without an explicit sort nor worker completion order determines row order.
Build item 7 tests multiple partitions and segments with shuffled decode
completion and verifies identical rows and worksheet splits.

For name collisions and the reserved-name check, compare the final candidate's
ECMAScript `String.prototype.toLowerCase()` value, without a locale argument or
locale-sensitive comparison. Do not use `toLocaleLowerCase()` or a collator. The
displayed name keeps its source case. Fixtures include `I` followed by `i` under
English and Turkish browser locales; both must allocate `I`, then `i_2`.

Name allocation reserves suffix space before truncation, as the existing
workbook merge allocator does. Replace forbidden worksheet-name characters with
`_`, remove leading and trailing whitespace and apostrophes, and use `Sheet` for
an empty base. The first part has no part suffix; later parts use `_2`, `_3`,
and so on. If a candidate is already claimed, append a collision suffix starting
at `_2` and increment until free. For each attempt, concatenate the part and
collision suffixes, shorten the sanitized base to leave room for that entire
suffix within 31 UTF-16 code units, then append it. Keep the longest prefix that
does not split a surrogate pair, and remove trailing whitespace or apostrophes
exposed by truncation. Treat the reserved name `History` as already claimed.
Each attempt starts from the original sanitized base, not a previously shortened
candidate. For example, a 31-character base leaves 29 characters before `_2`, 28
before `_10`, and 27 before `_2_2`. These constraints follow
[Excel's worksheet-name rules](https://support.microsoft.com/en-us/excel/rename-a-worksheet).

## Decision 6: value formatting

Boolean values are native Excel Boolean cells: `true` remains `true`, `false`
remains `false`, and null is a blank cell. They are not numeric `0` or `1`,
text, or formulas. Boolean columns preserve their source row positions;
independent workbook round trips must distinguish both Boolean values from null.
This unchanged representation adds no value-conversion manifest entry.

**Decision: write a finite numeric value as an Excel number only when it passes
all three checks: at most fifteen significant decimal digits, the numeric range
below, and an exact source-value round trip after workbook serialization. If any
check fails, write lossless decimal text and count `PBI_NUMERIC_AS_TEXT`. Dates
follow the separate rounding and fallback rules below. Power BI display format
strings are ignored in the first version.**

Use a conservative numeric-cell range based on
[Excel's published limits](https://support.microsoft.com/en-us/excel/excel-specifications-and-limits):
positive zero, or an absolute value from `2.2251e-308` through
`9.99999999999999e307`, inclusive. Thus `1e308` and `1e-310`, and their negative
counterparts, become text even when their decimal spelling is short. Preserve
negative zero as the text `-0`, also counted as `PBI_NUMERIC_AS_TEXT`.

An Excel number format changes display, not stored precision. Keep whole numbers
as their stored 64-bit integers and currency as its stored integer count of
ten-thousandths while forming the exact decimal text. For finite doubles, use
the locale-independent shortest decimal spelling that reconstructs the same
IEEE-754 value. Count significant digits in the coefficient, excluding leading
and trailing zeros and excluding the exponent. For example, `1234567890123456`
and the currency value `123456789012.3456` exceed the fifteen-digit limit and
become text. A value that passes the precision and range checks still must
survive serialization and reading back: compare whole numbers and scaled
currency in their original integer units, and doubles by their original finite
value. If it fails, use the prepared text instead. Currency numeric cells use a
four-decimal format. Build item 7 covers both numeric-range boundaries,
high-precision integers and currency, finite doubles, and both signs of zero;
build item 8 verifies the emitted cells with an independent reader.

Non-finite doubles are text: `NaN`, `Infinity`, and `-Infinity`, with their
affected-value count recorded as `PBI_NONFINITE_AS_TEXT`. These spellings
preserve the numeric category and infinity's sign; NaN payload bits are not
preserved. They are neither blank cells nor Excel errors. Null remains blank and
is not counted as a numeric conversion.

Dates are the exception because no exact representation exists: the model stores
them to the nanosecond and an Excel serial date is itself a double count of
days, so they are rounded to the millisecond by stated policy (the spike's only
cell differences were sub-millisecond remainders) and written as serials with a
date number format. Round to the nearest millisecond, resolving an exact half
millisecond toward the later instant. Count only emitted, non-null dates whose
stored timestamp changes under that rounding as `PBI_DATE_ROUNDED`; dates
already on a millisecond boundary do not increment it.

Use the workbook's 1900 date system. Test the original timestamp first: anything
before 1900-01-01 is written as text without rounding, even if rounding would
carry it into that day. For other timestamps, form the rounded candidate and
write a date serial only if that candidate is valid in the 1900 system and the
serialized double reads back to that same millisecond through the workbook
reader. A candidate in year 10000 or any failed round trip also selects text.
Every date-text fallback is derived from the original decoded timestamp, before
millisecond rounding, using the proleptic Gregorian calendar and the canonical
form `YYYY-MM-DDTHH:mm:ss.nnnnnnnnn`. Use two digits for each month, day, hour,
minute, and second, and exactly nine fractional digits, including trailing
zeros, from the decoded nanosecond units. Years 0000 through 9999 use four
digits; other years use a mandatory sign followed by the absolute year's decimal
digits, padded to a minimum width of six with no additional leading zeros. Thus
year 10000 is `+010000`, and year -1 is `-000001`. For example, one tenth of a
second is `.100000000`, and a whole second ends in `.000000000`. This spelling
does not depend on an original text representation or host locale. Do not attach
a timezone that the source does not carry. Count these cells as
`PBI_DATE_AS_TEXT` instead of `PBI_DATE_ROUNDED`; the fallback keeps the
original timestamp rather than the rounded candidate. The writer must not
generate Excel's fictitious 1900-02-29. Build-item 7 fixtures cover both
date-system limits, `1899-12-31T23:59:59.9996` remaining text, canonical
fractional zeros and expanded years, and a serial whose double conversion
crosses a millisecond boundary.

Honouring display format strings would mean implementing Power BI's
format-string language, which is out of proportion for a first version.

Binary values use standard padded base64 text without line breaks. Null remains
a blank cell; an empty byte sequence encodes as an empty string. The manifest
records the encoding and number of non-null encoded values for each binary
column. If any encoded value exceeds 32,767 characters, exclude that entire
column with `PBI_BINARY_CELL_TOO_LONG`, retain the other columns, and report the
number of oversized values. Do not truncate base64 or emit a partial column. The
build must verify byte-for-byte recovery, null and empty values, and the
text-length boundary before binary support is declared to work.

## Decision 7: hidden tables

Power BI generates hidden date tables (`LocalDateTable_*`,
`DateTableTemplate_*`) for every date column; thirteen of the sixty-two sample
tables were these.

**Decision: skip tables the model marks hidden by default, with an option to
include them; hiding a column does not exclude it from export.** Hidden tables
are noise to every user, and the model's own hidden flag is the criterion, not a
name prefix, so tables a user hid deliberately are treated the same way. Hidden
columns hold real data and follow the same decoding and representation rules as
visible columns. The manifest lists what was skipped.

## Decision 8: calculated tables and columns

**Decision: calculated tables and calculated columns are exported as data, and
their DAX expressions are recorded in the manifest.** They decode like any other
data and are data as far as the user is concerned; the expression is kept for
reference, not evaluated.

## Decision 9: a manifest bounded by the model structure

**Decision: aggregate value changes by table, column, and reason. Do not retain
an entry or a copy of the source value for each affected cell.** Each aggregate
holds a stable reason code and the affected-value count. The manifest also
records excluded tables and columns with their reasons, ordered worksheet parts
and their source row ranges, and the DAX expressions from Decision 8.

Source row ranges use zero-based, half-open coordinates `[start, end)` in the
table's concatenated row sequence from Decision 5. They exclude worksheet
headers and never restart at a partition or segment boundary. Part `k`, counted
from zero, covers `[k * 1048575, min((k + 1) * 1048575, rowCount))`. For
1,048,576 source rows the ranges are `[0, 1048575)` and `[1048575, 1048576)`. A
header-only table has one part with `[0, 0)`; excluded tables have no output
parts. Each range length equals that part's data-row count, and adjacent parts
meet without gaps or overlap.

The manifest is a UTF-8 JSON companion artifact with media type
`application/json` and `schemaVersion: 1`. It is present for every successful
export, including one with no warnings. Build item 8 commits its public schema,
typed serializer, and deterministic property-order fixtures before exposing the
operation. The schema carries the ordered table, column, part, reason, and DAX
data defined here; it adds no wall-clock timestamps, random IDs, or source cell
values.

Use the existing `ByteOperationOutcome`: `outputs` contains the workbook first
and the manifest second, and `result.artifacts` lists the matching file names
and media types in that same order. The portable workbook name comes from the
output plan; replacing its final `.xlsx` extension with `.manifest.json` gives
the companion name. No manifest field is added to shared `OperationResult`, and
the manifest is not a workbook worksheet. The file adapter and later CLI plan
and validate both destinations before writing, including overwrite and input
collision checks, and report both files as artifacts. The browser exposes both
downloads and includes both in its combined download. Refusals return structured
errors instead of a manifest-only successful artifact pair.

The manifest uses this reason-code vocabulary:

| Code                                     | Scope  | Meaning                                                       |
| ---------------------------------------- | ------ | ------------------------------------------------------------- |
| `PBI_TABLE_HIDDEN`                       | Table  | Hidden table excluded by the default policy                   |
| `PBI_TABLE_TOO_WIDE`                     | Table  | Source table exceeds 16,384 columns                           |
| `PBI_TABLE_NO_EXPORTABLE_COLUMNS`        | Table  | No column remains decodable and representable                 |
| `PBI_TABLE_SPLIT`                        | Table  | Rows span several worksheets; entry lists parts and ranges    |
| `PBI_COLUMN_UNSUPPORTED_ENCODING`        | Column | Reader does not support the column's encoding                 |
| `PBI_COLUMN_UNREADABLE`                  | Column | Corrupt or truncated column data prevents complete decoding   |
| `PBI_COLUMN_ROW_ALIGNMENT_UNRECOVERABLE` | Column | Decoded values cannot be assigned reliably to source rows     |
| `PBI_BINARY_CELL_TOO_LONG`               | Column | Column excluded because a base64 value exceeds the text limit |
| `PBI_NUMERIC_AS_TEXT`                    | Values | Numeric values written as exact text under Decision 6         |
| `PBI_NONFINITE_AS_TEXT`                  | Values | Non-finite doubles written with the defined text spelling     |
| `PBI_DATE_ROUNDED`                       | Values | Dates changed by rounding to the nearest millisecond          |
| `PBI_DATE_AS_TEXT`                       | Values | Original timestamps written as text when date serials fail    |
| `PBI_BINARY_AS_BASE64`                   | Values | Non-null binary values written as base64 text                 |
| `PBI_TEXT_TRUNCATED`                     | Values | Ordinary text values shortened to the worksheet cell limit    |

For each table in catalog order, emit its Table-scope reasons in ascending ASCII
code order. Then visit its columns in catalog order, emitting each column's
Column-scope reasons followed by its Values-scope aggregates, each group sorted
by code in ascending ASCII order. Values scope means a count attached to that
column, not a separate per-cell entry or unordered list. Omit aggregates with a
zero count. Worksheet parts remain in source row order. DAX expressions and
worksheet provenance are metadata, not additional reason codes.

Apply table exclusions in order: hidden policy, source column limit, then no
exportable columns. A skipped column has one exclusion code: unsupported
encoding first, unreadable data next, unrecoverable alignment next, then the
binary text limit. Later checks apply only after earlier checks succeed. Keep
the column exclusions when they cause `PBI_TABLE_NO_EXPORTABLE_COLUMNS` so the
user can see why no columns remain. Value-change counts describe emitted cells
only; discard conversion counts for an excluded column. The binary-limit entry
instead counts its oversized values. A `PBI_NO_EXPORTABLE_TABLES` error returns
these same table and column exclusions in its structured details.

This replaces per-value reporting: a million high-precision identifiers add one
count for their column, not a million manifest entries. Counts cover exact
numeric text, non-finite text, date rounding, date text, base64 conversion, and
ordinary-text truncation separately. The manifest does not duplicate those cell
values or collect row-index lists. Its size still depends on table and column
counts, worksheet parts, and DAX text, so its full size remains part of the
browser memory measurement.

The operation's `warnings` are a deterministic summary derived from the
manifest. Emit one plain-language warning per distinct reason code present,
sorted by code in ascending ASCII order, aggregating affected tables, columns,
or values in the unit that code describes. Include exclusions, worksheet splits,
text conversions, date rounding, and truncation. Each warning explains what
happened and directs the user to the manifest for detail; hidden-table warnings
also name the include-hidden option. Do not include source cell values or
construct per-cell warning strings. DAX and provenance metadata alone do not
produce warnings. Build item 8 verifies that a successful partial export has
these warnings and that the normal result renderer displays them; it must not
announce that no recoverable problems occurred after data was omitted or
changed.

## Known limits

The spike's corpus never exercised these paths. The code for them exists and
follows the reference implementation, but has not run on real input, so the
first build must treat each as unverified until a fixture is found:

- multithreaded XPress9 streams, which larger models produce;
- columns with several partitions, or several segments (models past roughly
  eight million rows);
- boolean and binary column types;
- string dictionaries in non-Latin scripts;
- compression classes other than hybrid run-length encoding;
- encrypted or password-protected models, which are refused, not read.

The pipeline was verified in Chromium only. Nothing in it is Chromium-specific,
but that is reasoning, not evidence. The spike measured only the decoder's
working set: it tracks the largest single table, not the whole model, and the
naive representation peaked at 284 MB on a 383,000-row model, with a
column-codes-and-dictionary representation expected to be several times lower.
That is not the whole bound. The browser surface must also hold the finished
workbook bytes, which accumulate across every exported worksheet, and the writer
materializes the workbook while producing them, so a model of many moderate
tables can exhaust a tab well below what the decoder figure implies. The build
must measure and bound decoder working set, workbook, zip, and manifest memory
together, including final serialization. The corpus must include many moderate
tables and a column with millions of values requiring exact text, to verify that
manifest entries grow with columns and reasons rather than row count. This is
required before the browser surface is declared to work, together with tests
that over-limit inputs receive `PBI_EXPORT_LIMIT_EXCEEDED` before the prohibited
allocation or destination write. These tested limits define the supported
envelope; they cannot guarantee spare memory in an already exhausted browser
process.

## Build list

Right-sized pull requests, in order, each with a contract agreed before code and
an independent review before push:

1. Package skeleton, zip and part reader, refusal contract
   (`PBI_INVALID_OPTIONS`, `PBI_INVALID_CONTAINER`, `PBI_MODEL_UNREADABLE`,
   `PBI_NO_MODEL`, `PBI_MODEL_ENCRYPTED`, `PBI_NO_EXPORTABLE_TABLES`,
   `PBI_EXPORT_LIMIT_EXCEEDED`) and error codes. Ships a real, testable refusal
   before any decode.
2. Vendored XPress9 source, Emscripten build script, committed binary, the
   same-origin copy script, and the licence attributions.
3. XPress9 chunk framing (single and multithreaded), the backup container, and a
   golden hash test per fixture.
4. Catalog over sql.js: storage schema, legacy schema fallbacks, the table and
   column model, and the hidden-object rule.
5. VertiPaq decoding: segment metadata, run-length and bit-packed columns,
   numeric dictionaries, value encoding, nulls.
6. String dictionaries: the Huffman kernel port and all three page encodings.
7. Typed values: integers, doubles, dates, currency, booleans, binary; the
   formatting policy including lossless base64 and whole-column exclusion for
   oversized binary cells; concatenation across partitions and segments.
8. Workbook emission through `@consultchimps/xlsx`: limits, worksheet splitting,
   the aggregated manifest, deterministic bytes, a corpus test.
9. The browser tool page, worker wiring, progress, cancellation, and the
   registry entry with its category.

## Consequences

- The registry gains a fourth category and one operation; the README table and
  drift checks change with it. ADR 0001's rules are unchanged.
- The repository gains its first vendored C source and committed WebAssembly
  binary, with a reproducibility check in CI. All licences are MIT.
- Output is faithful in substance, not in presentation: the workbook carries the
  model's values, not Power BI's formatting, and numeric date cells lose
  sub-millisecond detail by stated policy.
- The library and CLI surfaces share the same bytes-level engine, so the
  operation can ship on all three surfaces; the browser surface is first.
- Reading query definitions from a `.pbix` is not planned. Files that need it
  are old or are templates, and the rows path serves current files.
