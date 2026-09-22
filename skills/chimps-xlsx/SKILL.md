---
name: chimps-xlsx
description:
  Combine, split and inspect existing Excel workbooks with the ConsultChimps
  CLI, including choosing the right operation, resolving header rows under title
  blocks, and folding mismatched headers from many suppliers onto one schema
  with a column mapping. Use when .xlsx or .xlsm files that already exist must
  be stacked, tabbed together, divided by a column, or described before a
  decision. For authoring a new workbook as a deliverable, use
  chimps-excel-design-skill instead.
license: MIT
metadata:
  cli-version: "0.12.0"
  repository: consultchimps/consultchimps
---

# Excel operations with ConsultChimps

Transforming workbooks that already exist. Authoring a new one is
`chimps-excel-design-skill`; invocation, the JSON envelope and exit codes are
`use-consultchimps`.

Run everything as `npx consultchimps@0.12.0 <command>`.

## Inspect before you decide

Never choose an operation from a filename. Inspection writes nothing and tells
you what an operation will key on:

```bash
npx consultchimps@0.12.0 sheets inspect submissions/vendor-a.xlsx
```

It reports each worksheet with its visibility and used range, the header row
that will be resolved, the column names read from that row, and a few stored
values per column. Read three things from it: whether the header row is the one
you expect, whether the headers are spelled the way the other files spell them,
and whether data sits on a hidden worksheet.

## Choosing the operation

| The user wants                                      | Operation            |
| --------------------------------------------------- | -------------------- |
| One table of all rows, columns matched by header    | `sheets consolidate` |
| One workbook, each source sheet kept as its own tab | `sheets merge`       |
| One workbook per distinct value of a column         | `sheets split`       |
| To know what is in a file                           | `sheets inspect`     |

"Combine" and "merge" in a user's sentence usually mean consolidate. Confirm the
shape before running: stacked rows, or separate tabs.

## Recipe: submissions whose headers disagree

Several suppliers return the same form with different column names. Four steps,
and the middle one is a human review that must not be skipped.

```bash
# 1. Read the spellings each file actually carries
npx consultchimps@0.12.0 sheets inspect submissions/vendor-a.xlsx

# 2. Draft a mapping from the headers found, and the workbook alongside it
npx consultchimps@0.12.0 sheets consolidate "submissions/*.xlsx" \
  --suggest-map draft-mapping.json \
  -o consolidated.xlsx

# 3. Review and edit draft-mapping.json by hand. See references/mapping.md.

# 4. Apply the reviewed mapping
npx consultchimps@0.12.0 sheets consolidate "submissions/*.xlsx" \
  --map mapping.json \
  -o consolidated.xlsx
```

A draft groups only headers that already normalize to the same key, such as
`Failed Checks` and `Failed_Checks`. It never proposes that `Timestamp` and
`Run Time` are one column, because nothing in the data says they are. Adding
those is the reviewer's job, which is why step 3 exists.

`--map` and `--suggest-map` cannot be combined. Run twice if you want both.

When the only difference is spacing, case or punctuation, `--normalize-headers`
handles it without a mapping file at all. Reach for a mapping when the words
differ.

Every output row carries `_source_file`, `_source_sheet` and `_source_row`
unless you pass `--no-source`. Keep them: they are how a client traces a number
back to the file it came from.

## Recipe: split a master by column

```bash
npx consultchimps@0.12.0 sheets split master.xlsx \
  --column Region \
  --values \
  -o by-region
```

`--values` replaces formulas with their stored results, which is what you want
when the outputs go to people who should not see or break the model behind them.
Without it, formulas travel and may point at rows that are no longer there.

`--strict` matches values exactly, including case and type. Without it, `100`
and `"100"` group together, which is usually right for messy data and wrong for
codes that only look numeric.

## Header rows under title blocks

In version 0.12.0 the first row holding any value is taken as the header row. A
report title, a banner or a "Prepared by" line above the real headers therefore
becomes the header row, and the real headers become the first data row. Two
symptoms: column names that read like a sentence, and one row missing from the
output.

Confirm with `sheets inspect`, then name the real row:

```bash
npx consultchimps@0.12.0 sheets consolidate "submissions/*.xlsx" \
  --header-row 4 \
  -o consolidated.xlsx
```

`--header-row` is one-based and matches the row number Excel shows. It applies
to every file in the run, so files whose title blocks differ in height need
separate runs.

## Refusals worth recognising

| Code                          | Meaning and fix                                                                                             |
| ----------------------------- | ----------------------------------------------------------------------------------------------------------- |
| `XLSX_SPLIT_COLUMN_NOT_FOUND` | the column name is not on the resolved header row; inspect, then use `--header-row` or correct the spelling |
| `XLSX_INVALID_HEADER_ROW`     | the row named by `--header-row` is outside the sheet's used range                                           |
| `XLSX_EMPTY_HEADER`           | a column in the read range has no header text                                                               |
| `XLSX_DUPLICATE_HEADER`       | two columns on the header row carry the same name                                                           |
| `XLSX_MAPPING_FILE_INVALID`   | the mapping document is not valid JSON or breaks its own rules                                              |

A refusal means nothing was written. Fix the cause rather than retrying or
falling back to a script.

## Reference

- [Column mapping document format](references/mapping.md)
- [Every command and option in 0.12.0](../use-consultchimps/references/cli-reference.md)
