---
name: chimps-xlsx
description:
  Consolidate, merge, split and inspect existing Excel workbooks with the
  ConsultChimps CLI. Covers choosing the operation, header rows under title
  blocks, folding mismatched supplier headers onto one schema with a column
  mapping, and reconciling row counts. Use when .xlsx or .xlsm files that
  already exist must be stacked, tabbed together, divided by a column, or
  described before a decision.
license: Apache-2.0
metadata:
  cli-version: "0.12.0"
  repository: consultchimps/consultchimps
---

# Excel operations with ConsultChimps

Transforming workbooks that already exist. Run everything as
`npx consultchimps@0.12.0 <command>`; add `--json` before the command for a
machine-readable result. Every flag used here is in
[references/cli-reference.md](references/cli-reference.md).

## Choosing the operation

| The user wants                                      | Operation            |
| --------------------------------------------------- | -------------------- |
| One table of all rows, columns matched by header    | `sheets consolidate` |
| One workbook, each source sheet kept as its own tab | `sheets merge`       |
| One workbook per distinct value of a column         | `sheets split`       |
| To know what is in a file                           | `sheets inspect`     |

"Combine" and "merge" in a user's sentence usually mean consolidate. Confirm the
shape before running: stacked rows, or separate tabs.

## Inspect before you decide

Inspection writes nothing and tells you what an operation will key on:

```bash
npx consultchimps@0.12.0 --json sheets inspect submissions/vendor-a.xlsx
```

Per worksheet it reports visibility, `headerRow`, the column names read from
that row, sample values, `rowCount` (the used range) and `dataRowCount` (rows
that hold a value). Check three things: the header row is the one you expect,
the headers are spelled the way the other files spell them, and no data sits on
a hidden worksheet.

## What the reader does for you, and what it does not

Behaviour of 0.12.0, which a checkout build can differ from.

| In the source                               | What happens                                        |
| ------------------------------------------- | --------------------------------------------------- |
| Blank rows between or after the data        | skipped; the gap shows as `rowCount > dataRowCount` |
| Hidden worksheets                           | skipped unless you pass `--hidden`                  |
| Title or banner rows above the header       | taken as the header row: see below                  |
| Empty spacer columns                        | kept, named `column_N`, empty                       |
| Duplicate header names                      | renamed `Name_2`, with no warning                   |
| Blank header over a filled column           | named `column_N`, with no warning                   |
| "Mandatory" / "Optional" rows under headers | kept as data rows                                   |

The last four change your output without a warning. After a run, list the output
columns and look for `column_N` and `_2` names, and filter the first rows of
each source for annotation text. Remove annotation rows in the source or after
the run, and say which you did.

## Header rows under title blocks

In 0.12.0 the first row holding any value is taken as the header row. A report
title above the real headers becomes the header row, and the real headers become
the first data row. Symptoms: a column named like a sentence, `column_N` names,
and one extra row per file.

Confirm with `sheets inspect`, then name the real row:

```bash
npx consultchimps@0.12.0 sheets consolidate "submissions/*.xlsx" \
  --header-row 4 \
  -o consolidated.xlsx
```

`--header-row` is one-based, matches the row number Excel shows, and applies to
every file in the run. Files whose title blocks differ in height need one run
per height. Combine those outputs with `--no-source`: they already carry
`_source_*` columns naming the original files, and without the flag the run
refuses with `TABLE_SOURCE_COLUMN_COLLISION`.

## Recipe: submissions whose headers disagree

```bash
# 1. Read the spellings each file carries
npx consultchimps@0.12.0 --json sheets inspect submissions/vendor-a.xlsx

# 2. Trial run: draft a mapping, and a workbook to read the columns from
npx consultchimps@0.12.0 sheets consolidate "submissions/*.xlsx" \
  --suggest-map draft-mapping.json -o trial.xlsx

# 3. Review and edit the draft by hand. See references/mapping.md.

# 4. Apply it
npx consultchimps@0.12.0 sheets consolidate "submissions/*.xlsx" \
  --map mapping.json -o consolidated.xlsx
```

A draft groups only headers that normalize to the same key, such as
`Failed Checks` and `Failed_Checks`. It never proposes that `Timestamp` and
`Run Time` are one column. When only spacing, case or punctuation differ,
`--normalize-headers` does the job without a mapping file.

**An empty draft does not mean nothing needs mapping.** It means no two headers
differ only in spelling. Find synonyms from the trial workbook: for each column,
list the `_source_file` values of the rows that fill it. Two columns filled by
different files, and never by the same file, are candidates, such as `#` and
`S.No.`. Confirm each from the sample values before adding it as an alias.

**Aliases apply to every input.** If one supplier uses a header for something
else, folding it moves that supplier's values too. Run that supplier separately,
or leave the column unfolded and say so.

`--map` and `--suggest-map` cannot be combined in one run. A mapping that names
only the folds reports every other column in a warning, which is loud but loses
nothing. To keep the warning for columns you did not expect, list every known
column as a canonical entry; the trial workbook's header row gives you the list.

## Reconcile before you hand over

Every output row carries `_source_file`, `_source_sheet` and `_source_row`
unless you pass `--no-source`. Keep them: they are how a client traces a number
back to its file.

1. Sum `dataRowCount` from `inspect --json` across the inputs.
2. Compare with `metrics.outputRows` of the consolidate result. They should
   match; annotation rows you removed afterwards explain a shortfall.
3. Where `rowCount` exceeds `dataRowCount`, confirm the gap is blank formatted
   rows rather than data past a gap.
4. Report the totals, and every warning, to the user.

A full walk through for many suppliers and several templates is in
[references/supplier-batches.md](references/supplier-batches.md).

## Recipe: split a master by column

```bash
npx consultchimps@0.12.0 sheets split master.xlsx \
  --column Region --values -o by-region
```

`--values` replaces formulas with their stored results, which is what you want
when the outputs go to people who should not see the model behind them.
`--strict` matches values exactly, including case and type. Without it, `100`
and `"100"` group together, which is right for messy data and wrong for codes
that only look numeric.

## Refusals and errors worth recognising

| Code                          | Meaning and fix                                                                                             |
| ----------------------------- | ----------------------------------------------------------------------------------------------------------- |
| `XLSX_SPLIT_COLUMN_NOT_FOUND` | the column name is not on the resolved header row; inspect, then use `--header-row` or correct the spelling |
| `XLSX_INVALID_HEADER_ROW`     | the row named by `--header-row` is outside the sheet's used range                                           |
| `XLSX_MAPPING_FILE_INVALID`   | the mapping file is not valid JSON                                                                          |
| `TABLE_MAPPING_INVALID`       | valid JSON that breaks a mapping rule; see [references/mapping.md](references/mapping.md)                   |
| `TABLE_MAPPING_*`             | other mapping failures, such as a column collision; see [references/mapping.md](references/mapping.md)      |

A refusal means nothing was written. Fix the cause rather than retrying or
falling back to a script.

`"code": null` with `Maximum call stack size exceeded` from `sheets inspect`
means a large worksheet without a stored used range. `sheets consolidate` reads
the same file through a different reader and succeeds: skip the inspect for that
file and count its rows from the consolidate output instead.

## Reference

- [Column mapping document format](references/mapping.md)
- [Many suppliers, several templates](references/supplier-batches.md)
- [The `sheets` commands and options of 0.12.0](references/cli-reference.md)
