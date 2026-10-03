---
"@consultchimps/db": patch
"consultchimps": patch
---

The native database engines now load when a database opens, not when
`@consultchimps/db/node` is imported, so the portable single-file CLI no longer
loads `better-sqlite3` and `@duckdb/node-api` for `--version`, `--help`, or any
command that opens no database. An engine that cannot load is now reported when
a database of that format opens, as `DB_NATIVE_ENGINE_UNAVAILABLE` rather than
as a problem with the file.
