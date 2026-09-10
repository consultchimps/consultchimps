---
"@consultchimps/db": minor
---

Add `Database.updateRecord(table, recordId, values)`, the write path for editing
an existing record. It finds the record by its Record ID, refuses an unknown
table, record, or column, refuses any attempt to change the Record ID (which is
assigned once), and converts every value through the same coercion an insert
uses, so a value that does not fit the column's declared type is refused with a
stable, namespaced code rather than stored in a shape the read side would later
call corrupt. It returns the Record ID it wrote and the values now stored, read
back through their declared types, so a caller can show what the database kept
rather than what it sent.

`Database.readRecords` now accepts options: `columns` reads only the named
columns (the Record ID is always read first, and names match like every other
identifier) and `limit` reads at most that many records in storage order. An
unknown or repeated column name, or a limit that is not a whole number within
the safe integer range, is refused with a stable error. The options type is
exported as `ReadRecordsOptions`.
