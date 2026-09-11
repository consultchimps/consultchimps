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
no exportable table. Any readable model with at least one exportable table
produces a workbook, and every excluded table or column is reported in a
manifest.** Before policy applies, input that cannot be read at all is refused
with a validation code rather than a leaked parser error: a file that is not a
zip or lacks the parts a `.pbix` must have (`PBI_INVALID_CONTAINER`), and a
model part whose compressed stream, backup container, or catalog is truncated or
corrupt (`PBI_MODEL_UNREADABLE`). A valid Power BI container without a
`DataModel` part is refused with `PBI_NO_MODEL`. A missing `DataModel` is not an
invalid container. The message says the file has no embedded model to export and
asks for a `.pbix` saved with imported data. Templates, live connections, and
DirectQuery-only files are possible causes, not diagnoses inferred from an
extension or a missing part. They share this code until fixtures establish
reliable distinctions. The presence of a `Connections` part is not a
live-connection signal: a current Microsoft sample carries one alongside a full
model, and the spike hit that false positive. An encrypted or password-protected
model is refused (`PBI_MODEL_ENCRYPTED`); the reader does not attempt it.

When a file decodes partially, for example one column uses an encoding the
reader has not seen, the operation still produces the workbook and lists each
skipped table or column in the manifest with a stable code. A table remains
exportable when at least one column can be decoded and represented. Retained
columns keep their row positions and nulls; dropping a column never shifts or
drops rows. A column whose complete values or row alignment cannot be recovered
is skipped as a whole. If no columns remain, the table is excluded. A table with
zero rows and at least one output column produces a header-only worksheet.

A workbook is only produced when at least one table is exportable. When the
model holds tables but every one of them is excluded, whether hidden under the
default policy, over the column limit, or without any usable output columns, the
operation fails with a stable code (`PBI_NO_EXPORTABLE_TABLES`) whose message
lists each table with the reason it was excluded, and names the include-hidden
option when that is the cure. An empty workbook is never a success.

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
table (such a model is pathological), and ordinary cell text above 32,767
characters is truncated and counted in the manifest. Binary encodings follow
Decision 6 instead and are never truncated. Worksheet names follow Excel's rules
(31 characters, forbidden characters replaced, case-insensitive uniqueness with
numeric suffixes).

Worksheet allocation is deterministic. Tables are processed in the model's own
catalog order (the order the file stores them), never in decode-completion
order, and each table claims all of its worksheet names before the next table is
considered: its sanitized name first, then its numbered parts. A name already
claimed, case-insensitively, whether by an earlier table's name or one of its
parts, takes the next free numeric suffix. So a real table named `Sales_2` that
follows a split `Sales` becomes `Sales_2_2`, and the same file always yields the
same names, tab order, and bytes. Within a table, columns follow catalog order
and rows follow storage order.

## Decision 6: value formatting

**Decision: a numeric value is written as an Excel number only when Excel will
read back exactly the stored value, which means its exact decimal form has at
most fifteen significant digits; otherwise it is written as exact decimal text
and counted in the manifest. Dates are the one stated exception, rounded to the
millisecond. Power BI display format strings are ignored in the first version.**

Two limits stack. An Excel numeric cell is an IEEE-754 double, so a number
format changes only the display, never the stored precision; and Excel itself
keeps only fifteen significant decimal digits of a number it reads, so a value
the double carries exactly can still lose its trailing digit on open
(`1234567890123456` becomes `1234567890123450`). The fifteen-digit test
therefore subsumes the double test, and it is one test for every numeric type,
not a per-type range: whole numbers are 64-bit integers; currency is a 64-bit
count of ten-thousandths, whose decimal form is the integer with the point moved
four places (`123456789012.3456` has sixteen significant digits and is written
as text). The writer keeps the stored integer as an integer, forms the exact
decimal text, and counts its significant digits; at most fifteen is written as a
number (currency with a four-decimal format), and anything else is written as
that exact decimal text and counted under its column in the manifest, so the
workbook never silently alters a value.

Dates are the exception because no exact representation exists: the model stores
them to the nanosecond and an Excel serial date is itself a double count of
days, so they are rounded to the millisecond by stated policy (the spike's only
cell differences were sub-millisecond remainders) and written as serials with a
date number format. Honouring display format strings would mean implementing
Power BI's format-string language, which is out of proportion for a first
version.

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
| `PBI_BINARY_AS_BASE64`                   | Values | Non-null binary values written as base64 text                 |
| `PBI_TEXT_TRUNCATED`                     | Values | Ordinary text values shortened to the worksheet cell limit    |

Table entries precede column entries within each table. Tables and columns
follow catalog order; reasons within either scope sort by code in ascending
ASCII order. Worksheet parts remain in source row order. DAX expressions and
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
numeric text, base64 conversion, and ordinary-text truncation separately. The
manifest does not duplicate those cell values or collect row-index lists. Its
size still depends on table and column counts, worksheet parts, and DAX text, so
its full size remains part of the browser memory measurement.

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
required before the browser surface is declared to work.

## Build list

Right-sized pull requests, in order, each with a contract agreed before code and
an independent review before push:

1. Package skeleton, zip and part reader, refusal contract
   (`PBI_INVALID_CONTAINER`, `PBI_MODEL_UNREADABLE`, `PBI_NO_MODEL`,
   `PBI_MODEL_ENCRYPTED`, `PBI_NO_EXPORTABLE_TABLES`) and error codes. Ships a
   real, testable refusal before any decode.
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
  model's values, not Power BI's formatting, and dates lose sub-millisecond
  detail by stated policy.
- The library and CLI surfaces share the same bytes-level engine, so the
  operation can ship on all three surfaces; the browser surface is first.
- Reading query definitions from a `.pbix` is not planned. Files that need it
  are old or are templates, and the rows path serves current files.
