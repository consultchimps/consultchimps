---
"@consultchimps/db": patch
"consultchimps": patch
---

The native database engines now load when a database opens, not when
`@consultchimps/db/node` is imported, so the portable single-file CLI no longer
loads `better-sqlite3` and `@duckdb/node-api` for `--version`, `--help`, or any
command that opens no database. A missing native engine is now reported when a
database of that format opens.
