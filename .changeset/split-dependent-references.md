---
"@consultchimps/xlsx": minor
---

Fix every reference that pointed at a row a workbook split moved, and make split
outputs byte-reproducible.

Splitting a workbook by column removes the rows that belong to other groups and
closes the gaps they leave. Until now only the rows and their cells were
renumbered: everything else that described those rows kept its original row
number, so a delivered workbook could highlight the wrong cells, validate the
wrong column, link from the wrong row, or double the wrong record.

All of it now moves with the rows, on both worksheet ranges and Excel Tables:

- merged cells,
- conditional-formatting and data-validation ranges,
- hyperlinks,
- cell comments, including the drawing anchor that positions the note,
- formulas, including shared and array formulas and the spans they claim,
- the worksheet's declared used range.

The split also writes every output through one deterministic package writer, so
splitting the same workbook twice now produces byte-identical files rather than
files whose contents match but whose timestamps do not. Parts the split does not
touch still travel through byte for byte.

Public API, error codes, metrics and warning shapes are unchanged.
