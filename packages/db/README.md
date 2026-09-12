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

For a saved Node.js plan, `prepareImportFile` from `@consultchimps/db/node`
combines creation, capture, and publication. Pass `path`, `database`, `sources`,
`recipe`, and `baselineRevision`, with optional `overwrite`,
`protectedInputPaths`, `signal`, and `onProgress`. It returns a plan reference
and an operation result containing the final file artifact. Reopen the file with
`openPreparedImport` to inspect or apply it. An existing plan remains unchanged
if capture, source verification, or cancellation fails before publication. List
filesystem-backed source, recipe, and context files in `protectedInputPaths` so
overwrite validation can reject those destinations.

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
business-record reconciliation is outside these operations.

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

See the
[database guide](https://consultchimps.github.io/consultchimps/docs/tools/data-workspace/)
and
[library guide](https://consultchimps.github.io/consultchimps/docs/libraries/).
