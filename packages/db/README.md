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

Captured source values and generated observation IDs remain distinct from vendor
identifiers and delivery events. Identical content can be reused while another
delivery records a new touch point. Changed files append observations; automatic
business-record reconciliation is outside these operations.

The former in-memory spike API and analytics grid have been retired. Browser
storage is an OPFS working copy; it does not synchronize to a selected OS file.

See the
[database guide](https://consultchimps.github.io/consultchimps/docs/tools/data-workspace/)
and
[library guide](https://consultchimps.github.io/consultchimps/docs/libraries/).
