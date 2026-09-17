# @consultchimps/pbi

Low-level Power BI container reading for ConsultChimps. This is the first
implementation slice of
[ADR 0004](../../docs/adr/0004-power-bi-table-export.md). Table decoding and
Excel export are not available yet. The package is private and unpublished until
the export ships.

## Read the model part

```ts
import { readPbiModelPart } from "@consultchimps/pbi";

// containerBytes is a Uint8Array supplied by the caller.
const modelBytes = readPbiModelPart(containerBytes, {
  inputBytes: 64 * 1024 * 1024,
  decodedBytes: 256 * 1024 * 1024,
  peakBytes: 768 * 1024 * 1024,
});
```

`readPbiModelPart` synchronously returns an independent `Uint8Array` containing
the `DataModel` ZIP part. It leaves the supplied bytes unchanged and performs no
file writes, network access, or runtime initialization. The returned bytes still
contain the model's compression and backup framing, not decoded rows.

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
inside the model stream needs the later backup reader. Nonempty model bytes with
a valid ZIP CRC are not proof that the model itself is readable.

ZIP-level DEFLATE on `DataModel`, ZIP64, and multi-volume archives are refused
with `PBI_EXPORT_LIMIT_EXCEEDED` and `reason: "unsupported-zip-layout"`. The
message names the layout, not a byte limit: the refusal exists because inflation
and ZIP64 sizes cannot be bounded before allocation. Every sample file examined
stores `DataModel` without ZIP compression. This restriction does not refer to
XPress9 compression inside a stored `DataModel` part.

## Reader limits

The example lists the defaults. Each supplied limit must be a positive safe
integer. Limits are inclusive. Invalid limits are collected in `inputBytes`,
`decodedBytes`, then `peakBytes` order before the input is accessed.

- `inputBytes`: the length of the supplied container view
- `decodedBytes`: the size of the returned model part, including STORE output
- `peakBytes`: an estimate covering the retained input backing buffer (its
  maximum size when the buffer is resizable), parsing structures, and the
  returned model buffer

The peak estimate reserves 64 KiB of scratch space, 4 KiB per directory entry,
and eight bytes per central-directory byte, in addition to the input backing
buffer and output buffer. These are conservative reader ceilings, not measured
browser export defaults. The later export pipeline must also account for
runtimes, decoded columns, workbook serialization, manifests, and browser
copies.

## Refusals

Expected failures throw `ConsultChimpsError`. Every `details` object carries a
`stage` of `options`, `container`, or `model-part`; limit refusals add the
option, its limit, and the required value. Messages and details omit supplied
names, model bytes, and raw parser errors. Checks on the model part run in the
order encryption flags, empty part, storage method, size consistency, byte
limits, then CRC, so an empty DEFLATE-stored part reports
`PBI_MODEL_UNREADABLE`.

| Code                        | Meaning                                                                         |
| --------------------------- | ------------------------------------------------------------------------------- |
| `PBI_INVALID_OPTIONS`       | One or more reader limits are invalid                                           |
| `PBI_INVALID_CONTAINER`     | ZIP structure or required container markers are invalid                         |
| `PBI_NO_MODEL`              | The container markers exist but `DataModel` is absent                           |
| `PBI_MODEL_ENCRYPTED`       | The model part has ZIP encryption flags                                         |
| `PBI_MODEL_UNREADABLE`      | The stored model part is empty, has inconsistent sizes, or fails CRC validation |
| `PBI_EXPORT_LIMIT_EXCEEDED` | A configured bound is exceeded or the ZIP layout lacks a supported bound        |

`PbiErrorCode` also reserves `PBI_RUNTIME_UNAVAILABLE` and
`PBI_NO_EXPORTABLE_TABLES` for later pipeline stages. The reader does not load
SQLite or XPress9, inspect table eligibility, or accept export/runtime options.
