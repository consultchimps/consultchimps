---
"@consultchimps/db": minor
---

Add the `@consultchimps/db` package: a local in-memory relational database built
on sql.js. It loads a database from bytes, runs queries and statements, and
serializes back to bytes through a thin engine wrapper; models tables, columns,
column types, and foreign keys; and generates a human-readable, immutable
per-table Record ID (for example `CUST-0001`) that foreign keys reference. The
schema, per-table identifier configuration, and next-id counter persist inside
the database file itself, so a reopened file restores its own schema and id
state. A `Table` bridge converts between the `@consultchimps/tabular` table
model and database tables in both directions.
