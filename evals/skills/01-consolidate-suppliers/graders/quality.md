---
type: llm
target: trace
---

Pass when all of these hold: it ran `sheets consolidate` from the ConsultChimps
CLI rather than writing a script; it noticed that supplier-c has two title rows
above its header (header on row 3) and handled them; SKU, Price and Qty from
supplier-d end up in the same columns as Item Code, Unit Price and Quantity; the
lower-case headers of supplier-b fold onto the others; the output has 54 data
rows and keeps the _source columns; the final message reports the row
reconciliation and any warnings.
