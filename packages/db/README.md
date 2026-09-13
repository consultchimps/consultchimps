# @consultchimps/db

Persistent local SQLite and DuckDB databases for ConsultChimps.

The root entry provides database schemas, import recipes, durable source
captures, and delivery operations. Runtime adapters use `@consultchimps/db/node`
and `@consultchimps/db/browser`. The CLI exposes them under `consultchimps db`.

```ts
import { inspectDatabase } from "@consultchimps/db";
import { createDatabase } from "@consultchimps/db/node";

const created = await createDatabase({
  path: "inventory.duckdb",
  format: "duckdb",
});
try {
  console.log(await inspectDatabase({ database: created.database }));
} finally {
  await created.database.close();
}
```

Node consumers install `@duckdb/node-api` and `better-sqlite3` alongside this
package. Private import plans use SQLite even when the working database is
DuckDB. Browser consumers provide the pinned `@duckdb/duckdb-wasm` and
`@sqlite.org/sqlite-wasm` runtimes and serve their assets from their own origin.
The runtime entries keep native bindings out of the browser operation layer.

Import composition:

1. Create or open a persistent working database.
2. Supply workbook sources through `createWorkbookImportSource`, or implement
   the `ImportSource` byte and bounded-row contracts.
3. Create a separate prepared-import handle and call `prepareImport`.
4. Review `inspectImport`, then use `resolveImport` for destination decisions.
5. Pass its ready revision to `applyImport` with a retry request ID.
6. Close handles and release scratch files.

Custom `ImportSource` readers must provide valid decimal or exponent numeric
tokens and valid date payloads. Date `iso` values must name real calendar dates
or UTC timestamps; textual `raw` values must agree with them. Numeric date
serials must be finite, and their reader owns epoch conversion. The same checks
apply to cached formula values and saved plan rows. Malformed values are
rejected instead of being coerced to zero or stored as invalid dates.

Numeric date tokens are not restricted to Excel epochs. A custom reader can
decode Unix milliseconds into `iso`, for example. Typed date mappings use that
decoded value; text mappings retain `raw`. The XLSX adapter performs its own
conversion using the workbook's declared date system. Captures do not retain
structured epoch metadata, so the DB does not independently recalculate numeric
dates or verify that an externally edited token and ISO value agree.

Use `inspectImport({ database, prepared, page: { limit: 20 } })` to preview both
newly captured rows and rows reused from the working database. Omitting
`database` still returns plan metadata and staged rows, with
`DB_PREVIEW_DATABASE_REQUIRED` entries in `previewWarnings` for reused
selections. Preview pages contain at most 100 rows; pass `nextCursor` as the
next page's `cursor`. The target database must match the plan's database ID.

Preparation metrics count input sources, not worksheet selections. `sourcesRead`
counts sources with a newly captured selection; `sourcesReused` counts sources
with a reused capture. A source with both kinds contributes once to each count.
Selections already bound in the same saved plan contribute to neither count on
retry. `rowsCaptured` counts rows captured during this preparation call.

For a saved Node.js plan, `prepareImportFile` from `@consultchimps/db/node`
combines creation, capture, and publication. Pass `path`, `database`, `sources`,
`recipe`, and `baselineRevision`, with optional `overwrite`,
`protectedInputPaths`, `signal`, and `onProgress`. It returns a plan reference
and an operation result containing the final file artifact. Reopen the file with
`openPreparedImport` to inspect or apply it. An existing plan remains unchanged
if capture, source verification, or cancellation fails before publication. List
filesystem-backed source, recipe, and context files in `protectedInputPaths` so
overwrite validation can reject those destinations.

Prepared plans use artifact format version 3. Their stored review fingerprint
binds the plan identity, database baseline, revision, state, recipe, conflicts,
decisions, source bindings, and capture definitions, including a checksum of
each capture's row coordinates and serialized values. Reading a plan checks its
metadata fingerprint, and applying requires the same `reviewFingerprint` as the
approved reference. Inspection and apply read capture definitions in the same
snapshot as that metadata. Preparing or resolving a review rejects a stale write
if the plan changed during evaluation; inspect and review its latest revision
before retrying.

Fresh capture checksums are computed during capture and verified during the
existing copy into the working database, including excluded selections. Reusing
a capture requires one bounded read of its stored rows during preparation. Apply
verifies reused rows when consuming them for a new table application. Checksum
mismatches roll back the transaction before commit. Receipt-only retries and
already-applied table applications do not rescan row contents.

These are integrity checks, not digital signatures. An editor who coherently
rewrites the artifact and its checksums can produce a different valid artifact.
Version 1 and 2 spike plans are unsupported. Regenerate them from their original
sources with this build. Existing working database files keep their format.

Use `openPreparedImport({ path, readonly: true })` for inspection. This opens
the saved plan without enabling SQLite's writable journal mode. Omit `readonly`
when preparing, resolving, or applying a plan, because those operations update
its saved state. The CLI uses read-only access for `db inspect`.

The Node runtime refuses to replace a database or plan held open through the
same runtime, including filesystem aliases, with `DB_NATIVE_FILE_BUSY`. Close
those handles before replacement. Callers must also prevent other processes from
opening or writing the destination during replacement.

Native exports check cancellation after copying and validation, and again before
entering file publication. Cancellation before that boundary leaves the
destination unchanged and attempts to remove the private staged copy.
Cancellation is no longer checked once file publication begins.

`consultchimps db resolve --recipe` replaces the saved plan's table routes. A
selection omitted from the replacement recipe is excluded from table loading.
Its captured rows remain in the plan and become received evidence in the managed
database when you apply the plan. A recorded delivery also retains that
selection. Library callers can use `replaceImportRecipe` for the same
replacement semantics, including stored recipe routes that have no captured
selection.

Captured source values and generated observation IDs remain distinct from vendor
identifiers and delivery events. Identical content can be reused while another
delivery records a new touch point. Changed files append observations; automatic
business-record reconciliation is outside these operations. Reusing a capture in
the same destination table requires the same effective column mapping. A changed
mapping returns `DB_IMPORT_APPLICATION_CONFLICT` before commit; choose a new
table to retain another interpretation without replacing earlier observations.

If two aliases for the same captured selection target one table with different
effective mappings, preparation and review report
`conflicting-application-mapping` for the affected routes. Use matching mappings
or separate destination tables before applying. Equivalent alias routes can
share one application.

`inspectImport` reports capture reuse separately from each route's
`applicationState`. The `already-applied` state requires a matching captured
source, destination table, and effective column mapping. Reusing captured Excel
content does not imply that loading it into another table adds no rows.

`listDeliveries` returns delivery pages in allocation order. Each record lists
the capture IDs that also appeared in an earlier delivery as `reusedCaptureIds`,
including when that earlier delivery is on another page.

Cross-format export checks stored values against the registered logical schema
as it copies bounded batches. External edits that violate those types stop the
conversion before publication instead of relying on the destination engine to
coerce them. For example, SQLite can store `2` in a Boolean column or `1.5` in
an integer column; those values need correction before conversion to DuckDB. An
export dry run checks schema compatibility without scanning stored rows.

A same-format export retains the logical database ID. A saved review can resume
against that snapshot when its revision and schema still match. Cross-format
conversion creates a new database ID and requires a new review. Copies do not
synchronize, and a shared ID is not a lock or a guarantee of equal row contents.

Browser cross-format export uses a temporary conversion working copy. If cleanup
fails after export succeeds, the result includes a warning identifying that
copy. If export and cleanup both fail, `DB_BROWSER_CONVERSION_CLEANUP_REQUIRED`
reports `conversionName`, `format`, and `directory`. Reconfigure the same
browser storage and use `openDatabase` with that logical name to inspect or
export recoverable data after resolving the storage issue. SQLite pool files
have opaque physical names; the reported conversion name is a logical database
name, not an OPFS filename.

The former in-memory spike API and analytics grid have been retired. Browser
storage is an OPFS working copy; it does not synchronize to a selected OS file.

Browser database create and import operations, and prepared-plan creation,
validate a temporary candidate before replacing an existing working copy. A
replacement can temporarily require space for the existing working copy, the
candidate, and a recovery backup. Closing the open handle is required before
replacement. Cancellation is checked before publication begins; after that
boundary the runtime finishes publication or restores the backup so the logical
name does not point at a partial working copy. If publication and restoration
both fail, the error reports the retained backup name and its kind. A library
caller can open a database backup with `BrowserDatabaseRuntime.openDatabase` and
export it, or open a prepared-plan backup with `openPreparedImport` to inspect
or resume it. This recovery covers failures reported to the running operation.
An abrupt tab, worker, or browser termination can interrupt publication. DuckDB
stores its main file and write-ahead log as separate OPFS entries, so browser
replacement is not crash-atomic across those entries.

`BrowserDatabaseRuntime.listPreparedImports` returns readable `imports` and
`ignored` entries with a `name`, stable error `code`, and recovery `message`.
Unsupported plan versions remain stored and are reported through those entries.

Call `BrowserDatabaseRuntime.discardPreparedImport({ name })` only for a private
plan that the caller created and no longer needs. The runtime refuses to remove
an open plan, a database, or an unrecognized artifact with that name. This
operation does not provide general browser database deletion.

Use `openDatabase({ name, readonly: true })` for browser inspection or export
without permitting database writes. Browser exports require an empty
`RandomAccessFile` destination unless `overwrite: true` is explicit. The caller
owns the destination and must keep its backing storage separate from the working
database.

A read-only DuckDB export copies the open database into a temporary OPFS
database using DuckDB's `COPY FROM DATABASE` operation. It checkpoints the
temporary database so the exported file includes committed changes visible to
the source connection without checkpointing the source. This requires additional
browser storage for the temporary copy. If cleanup fails,
`DB_BROWSER_DUCKDB_SNAPSHOT_CLEANUP_REQUIRED` reports its `directory` and
`snapshotName`. Release open handles before inspecting or removing the remaining
temporary main file and optional `.wal` file through OPFS. The original working
copy remains the source for a retry.

Export does not repair an incomplete recovery log. Committed writes in the
browser engine are not guaranteed to survive abrupt worker or browser
termination.

Replacing a nonempty export destination creates a temporary OPFS backup using
bounded reads. This requires additional browser storage for the prior output.
The runtime attempts to restore prior bytes and length after cancellation or a
reported write failure. If restoration fails,
`DB_BROWSER_EXPORT_RECOVERY_REQUIRED` identifies the retained file through
`backupDirectory` and `backupName` in its error details. Use
`navigator.storage.getDirectory()` and those names to retrieve that file before
retrying. Recovery of an initially empty destination requires clearing its
incomplete contents instead; no backup is needed for an empty file. Cleanup
failures produce a warning after successful export, or
`DB_BROWSER_EXPORT_CLEANUP_REQUIRED` after a failed export whose prior contents
were preserved. These recovery paths require the operation to remain running;
they do not make abrupt browser termination atomic.

See the
[database guide](https://consultchimps.github.io/consultchimps/docs/tools/data-workspace/)
and
[library guide](https://consultchimps.github.io/consultchimps/docs/libraries/).
