---
"@consultchimps/db": minor
---

Add `listDatabases` and `removeDatabase` to the browser database runtime.
`listDatabases` names the SQLite and DuckDB working copies stored in the browser
without opening any of them, leaving out private candidate, backup,
import-batch, and export files, and reports a name stored in both formats or a
DuckDB write-ahead log without its database as ignored with the code opening it
would raise. `removeDatabase` deletes a working copy's files under the same
per-name lock the other operations use and is refused while the copy is open.
The browser database tool uses both so its list of working copies reflects
storage rather than a remembered list, and so a copy can be deleted.
