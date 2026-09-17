# @consultchimps/pbi

Power BI model reading and table export for ConsultChimps. This is
[ADR 0004](../../docs/adr/0004-power-bi-table-export.md) build items 2 through
8: the package turns a `.pbix` into a workbook plus a JSON manifest in one call.
The CLI command, the registry entry, and the browser tool page are not here. The
package is private and unpublished until the CLI surface ships.

## Export every table

```ts
import { exportPbiTables } from "@consultchimps/pbi";

// containerBytes is a Uint8Array supplied by the caller.
const outcome = await exportPbiTables(containerBytes, {
  includeHiddenTables: false,
  outputName: "power-bi-tables.xlsx",
});

const [workbook, manifest] = outcome.outputs;
outcome.result.warnings.forEach((warning) => console.warn(warning));
```

`outputs` carries the workbook first and the manifest second, with the media
types `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet` and
`application/json`. `result.artifacts` lists the same two names and types in the
same order, and `result.metrics` always carries the six keys `inputFiles`,
`outputFiles`, `exportedTables`, `exportedColumns`, `exportedRows` and
`outputWorksheets`.

## Read the rows without a workbook

```ts
import { readPbiTables } from "@consultchimps/pbi";

const model = await readPbiTables(containerBytes);
for (const table of model.tables) {
  console.log(table.name, table.rowCount, table.columns.length);
}
```

`readPbiTables` returns decoded, typed values: `string`, `bigint` for `int64`
and for `currency` (the stored count of ten-thousandths, undivided), `number`
for `double` and for `dateTimeSerial` (the day serial from epoch 1899-12-30),
`boolean`, `Uint8Array` for binary, and `null` for a null cell. Every column's
`values` array has exactly `table.rowCount` entries, and `rowCount` is the
catalog's own declared count for the table, never a count inferred from what the
columns happened to decode to: a column whose decoded length differs from it is
excluded with `PBI_COLUMN_ROW_ALIGNMENT_UNRECOVERABLE`, and no row moves. Its
manifest records structure, exclusions and DAX but no value-conversion counts:
those describe emitted cells, and a decode produces none.

The package has three entry points. `.` carries the two functions above and
`readPbiModelPart`; `./xpress9.wasm` carries the committed decoder binary, so a
browser bundler can resolve the asset by package path and serve it from the
host's own origin rather than a CDN.

## Read the model part only

```ts
import { readPbiModelPart } from "@consultchimps/pbi";

const modelBytes = readPbiModelPart(containerBytes);
```

`readPbiModelPart` synchronously returns an independent `Uint8Array` containing
the `DataModel` ZIP part. It leaves the supplied bytes unchanged and performs no
file writes, network access, or runtime initialization. The returned bytes still
contain the model's compression and backup framing, not decoded rows.

## Runtimes

Two WebAssembly runtimes are involved: SQLite, for the model's storage catalog,
and XPress9, for the compressed backup stream. `runtime.sql` is passed straight
through to `@consultchimps/db/sqlite-read`; `runtime.xpress9` configures the
decoder this package ships.

| Asset          | Omitted configuration resolves to                                       |
| -------------- | ----------------------------------------------------------------------- |
| `sqlite3.wasm` | the installed `@sqlite.org/sqlite-wasm` peer dependency                 |
| `xpress9.wasm` | the copy committed beside this package, read through `node:fs/promises` |

A browser caller must configure both. For each configuration supplied, exactly
one source is required: a `locateFile` function that receives the asset name, or
a nonempty `wasmBinary`. Both, or neither, is `PBI_INVALID_OPTIONS` at that
child path. Configuration belongs to the invocation; a binary cached from a
different configuration is never reused.

Local verification of the committed binary is limited to comparing its SHA-256
with the value recorded in the build script. The two-clean-build reproducibility
check runs in CI inside the pinned Emscripten container; a passing local run is
not a reproducibility check.

## Supported container profile

- Single-volume, non-ZIP64 archives with UTF-8 part names
- Required part names: `[Content_Types].xml` and `Version`. Report parts are not
  required; older files carry `Report/Layout` and enhanced-format files carry
  `Report/definition/...`
- A nonempty `DataModel` part stored with ZIP method 0 (STORE)
- Matching central and local headers, non-overlapping entries, and optional ZIP
  data descriptors in either form, with or without the signature; the signed
  reading is tried first, so a CRC that equals the signature value still reads
- A matching CRC-32 for the model part, computed on the returned copy
- A live input view; a view whose buffer has been detached by a transfer is
  refused as an invalid container

The end-of-central-directory record is the highest-offset candidate whose
comment reaches the end of the input and whose directory ends exactly at the
record; a decoy signature inside an archive comment does not hide the real one,
even when the decoy carries ZIP64 or multi-disk fields. Bytes before the first
entry are ignored when the directory offsets are file-absolute, as
self-extracting writers produce them. Entries must not overlap, but they need
not cover every byte before the directory.

The other parts are not decompressed or semantically validated. A `Connections`
part does not imply that the file is a live connection. Missing `DataModel`
produces `PBI_NO_MODEL` regardless of the filename or connection parts.

ZIP encryption flags on `DataModel` produce `PBI_MODEL_ENCRYPTED`. Encryption
inside the model stream needs the backup reader to see it. Nonempty model bytes
with a valid ZIP CRC are not proof that the model itself is readable.

ZIP-level DEFLATE on `DataModel`, ZIP64, and multi-volume archives are refused
with `PBI_EXPORT_LIMIT_EXCEEDED` and `reason: "unsupported-zip-layout"`. The
message names the layout, not a byte limit: the refusal exists because inflation
and ZIP64 sizes cannot be bounded before allocation. Every sample file examined
stores `DataModel` without ZIP compression. This restriction does not refer to
XPress9 compression inside a stored `DataModel` part.

## What the workbook holds

One worksheet per table, split at 1,048,575 data rows per part, with the header
repeated on every part and the parts named `Sales`, `Sales_2`, `Sales_3`. Tables
appear in ascending catalog table id and columns in ascending storage position.
Names are sanitized, compared case-insensitively with no locale, and `History`
is treated as claimed.

Values follow ADR Decision 6. A finite number becomes a worksheet number only
when its coefficient has at most fifteen significant decimal digits, its
magnitude is positive zero or within Excel's published range, and it reads back
in its original units; anything else becomes lossless decimal text. Dates are
the decoded day serial, rounded once to an integer millisecond with an exact
half going to the later instant, written with a date format, and falling back to
`pbi-date-serial:` text outside years 1900 to 9999. Currency keeps its integer
count of ten-thousandths until formatting and uses a four-decimal format.
Booleans are native Boolean cells. Binary is standard padded base64. Text over
32,767 UTF-16 code units or 253 line breaks is truncated without splitting a
surrogate pair or a CRLF.

Serialization is lossless: every XML 1.0-forbidden code unit, every carriage
return, U+FFFE, U+FFFF and every unpaired surrogate is written as `_xHHHH_`, and
a literal `_xHHHH_` has its leading underscore encoded as `_x005F_`.

The workbook is byte-deterministic. Entry dates are fixed at 1980-01-01, the
compression method and level are fixed, entry order is fixed, and no comparison
anywhere is locale-sensitive.

## The manifest

A UTF-8 JSON companion with `schemaVersion: 1`, no wall-clock timestamp, no
random id and no copy of any source value. It lists every exported table with
its worksheet parts and their zero-based half-open source row ranges, every
excluded table and column with its reason code, the DAX of calculated tables and
columns (capped at 32,768 code units per expression), and the reference paths
that ran without corpus evidence.

Reasons are aggregated by table, column and code, never per cell: a million
high-precision identifiers add one count to their column. `result.warnings`
carries one plain-language line per distinct code present, in ascending ASCII
code order.

## Unverified paths

The spike corpus of nine public Microsoft samples never exercised these paths.
The code follows the reference implementation but has not run on real input, so
it attempts the decode and records a warning rather than refusing:

| Code                                   | Path                                           |
| -------------------------------------- | ---------------------------------------------- |
| `PBI_UNVERIFIED_XPRESS9_MULTITHREADED` | the multithreaded backup compression variant   |
| `PBI_UNVERIFIED_MULTIPLE_PARTITIONS`   | a column stored in several partitions          |
| `PBI_UNVERIFIED_MULTIPLE_SEGMENTS`     | a column stored in several segments            |
| `PBI_UNVERIFIED_NON_LATIN_DICTIONARY`  | a string dictionary in a non-Latin script      |
| `PBI_UNVERIFIED_BOOLEAN_TYPE`          | a true or false column                         |
| `PBI_UNVERIFIED_BINARY_TYPE`           | a binary column                                |
| `PBI_UNVERIFIED_COMPRESSION_CLASS`     | a column compression class the width map lacks |

A compression class with no known width and a bit-packed subsegment is the one
case that cannot attempt the decode: the column is excluded with
`PBI_COLUMN_UNSUPPORTED_ENCODING`.

Two more reading paths have no corpus evidence either: an uncompressed backup
image, and a catalog whose `Column` table lacks `Type` or `InferredDataType`, so
the `BindingType` and `ExplicitDataType` fallbacks take over. Every corpus
catalog carries the modern columns.

## Limits

Each supplied limit must be a positive safe integer and every limit is
inclusive. Invalid options are collected into one `PBI_INVALID_OPTIONS` in the
fixed order `inputBytes`, `decodedBytes`, `outputBytes`, `peakBytes`,
`includeHiddenTables`, `outputName`, `runtime`, `runtime.sql`,
`runtime.xpress9`, with at most one detail per path and never the supplied
value, before any input read or locator call.

| Option         | Default | Bytes counted                                                          |
| -------------- | ------- | ---------------------------------------------------------------------- |
| `inputBytes`   | 64 MiB  | the length of the supplied container view                              |
| `decodedBytes` | 256 MiB | every buffer decompression produces, counted when produced             |
| `outputBytes`  | 256 MiB | the workbook and the manifest together, reserved before either is sent |
| `peakBytes`    | 768 MiB | a bound over the buffers this reader allocates                         |

`peakBytes` bounds the buffers this reader allocates for the container, the
model part, the backup image and its header tree, the catalog, the two
WebAssembly runtimes, the decoded columns and their dictionaries, and the
worksheet XML with the finished package. Each is reserved before the allocation
it covers, sized from the catalog's declared row counts and the members' own
byte lengths, never by catching an out-of-memory.

It is not a bound on the process. A decoded cell is charged at its measured cost
by type, eight bytes for a boolean, sixteen for a dictionary string, twenty-four
for a double or a date serial and thirty-two for an int64 or a currency. Those
are the costs of the values themselves, not of everything the JavaScript engine
holds while it builds a workbook from them: a model whose largest table
approaches the row limit has been measured at about four times the reported
peak. Size the host for the measured figure, not for `peakBytes`.

`scripts/measure-cell-cost.ts` is where the per-type figures come from. Run it
again after any change to the decoded value representation.

These are Node figures measured on the corpus. They are not browser defaults:
the browser envelope, including page-held upload bytes, worker transport in both
directions, and download copies, is measured in the browser release.

## Refusals

Expected failures throw `ConsultChimpsError` with no cause, no upstream text and
no model-provided name, identifier, DAX or cell value. Named exclusions belong
only to a successful export's manifest.

| Code                        | Meaning                                                                  |
| --------------------------- | ------------------------------------------------------------------------ |
| `PBI_INVALID_OPTIONS`       | One or more options are invalid                                          |
| `PBI_INVALID_CONTAINER`     | ZIP structure or required container markers are invalid                  |
| `PBI_NO_MODEL`              | The container markers exist but `DataModel` is absent                    |
| `PBI_MODEL_ENCRYPTED`       | The model part has ZIP encryption flags                                  |
| `PBI_MODEL_UNREADABLE`      | The model stream, backup container, or catalog is truncated or damaged   |
| `PBI_NO_EXPORTABLE_TABLES`  | No table survived, with anonymous counts by exclusion code               |
| `PBI_RUNTIME_UNAVAILABLE`   | A WebAssembly runtime could not load, compile, or instantiate            |
| `PBI_EXPORT_LIMIT_EXCEEDED` | A configured bound is exceeded or the ZIP layout lacks a supported bound |

A catalog reader that cannot release itself stays a failure: it surfaces as
`PBI_MODEL_UNREADABLE` with `recovery: "discard-worker"` and never masks a
successful export.

## The fixture corpus

Two fixtures and their full cell-level dumps are committed, so the default test
run exercises the whole pipeline offline. The other seven are public Microsoft
samples pinned by repository commit in `fixtures/oracle-digests.json` and
fetched on demand:

```sh
node scripts/fetch-fixtures.ts
```

Every fetched file is verified against its committed SHA-256 before it is
written. The corpus test compares cell for cell against the committed dumps and
at per-table digest level against any fetched file present, and reports by name
every fixture it skipped.
