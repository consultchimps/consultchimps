---
"@consultchimps/db": minor
---

Add `readTableRows`, a bounded, read-only page of a managed table's stored rows
ordered by Record ID. The page carries the Record ID, the table's declared
columns in order, and the source file id, with 64-bit integers rendered as
decimal strings and binary values as a byte count. A `nextCursor` continues the
page; a cursor from another table, a malformed cursor, a page size outside 1 to
200, or an unknown table is refused with a stable code. The browser database
tool uses it for its row browser; nothing is written and no checkpoint is taken.
