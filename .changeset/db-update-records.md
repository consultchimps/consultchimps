---
"@consultchimps/db": minor
---

Add `updateRecords(database, requests)`, which applies many record updates as
one step. Every request runs inside a single transaction, so the accepted writes
commit together rather than landing one at a time, and a value the schema
refuses is reported against its own request while the writes beside it stand.
Each outcome carries either the stored values read back through their declared
types or the refusal's stable code and sentence. A failure that is not a refusal
rolls the whole batch back. It is the write path behind a spreadsheet-style
paste or fill, where one gesture is many cell writes.
