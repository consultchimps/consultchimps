# Many suppliers, several templates

The case this covers: a folder per template (for example a source list, a table
list and a column list), one workbook per supplier in each, and one consolidated
table wanted per template. Suppliers filled the same template differently: title
rows added, columns renamed, extra columns appended, serial number columns under
different names.

Run the steps below once per template folder. Keep a note of every decision; it
becomes the handover.

## 1. Inspect every input

In bash, zsh or Git Bash:

```bash
mkdir -p inspect
for f in templates/sources/*.xlsx; do
  npx consultchimps@0.12.0 --json sheets inspect "$f" > "inspect/$(basename "$f" .xlsx).json"
done
```

In Windows PowerShell:

```powershell
New-Item -ItemType Directory -Force inspect | Out-Null
Get-ChildItem templates/sources -Filter *.xlsx | ForEach-Object {
  npx consultchimps@0.12.0 --json sheets inspect $_.FullName |
    Set-Content -Encoding utf8 "inspect/$($_.BaseName).json"
}
```

From each result record the worksheet names and visibility, `headerRow`,
`rowCount`, `dataRowCount` and the header list. Look for:

- more than one visible worksheet: decide which one is the data, and pass
  `--sheet` with its exact name
- a header list that reads like a title or is mostly `column_N`: a title row is
  sitting above the header, so plan a `--header-row` run for that file
- `rowCount` much larger than `dataRowCount`: formatted blank rows, harmless,
  but confirm when the gap is large
- an `ok: false` result with `"code": null`: see the error notes in the skill

## 2. Group files by header row

`--header-row` applies to every file in a run, so each group of files that share
a header row needs its own runs. Put each group in its own folder, for example
`groups/row1` for files whose headers sit on row 1 and `groups/row3` for files
with a two row title block. Steps 3 and 4 run once per group folder, with that
group's `--header-row`; row 1 is the default and needs no flag.

The group outputs are then combined with `--no-source`. Each group output
already carries `_source_file`, `_source_sheet` and `_source_row` naming the
original workbooks, and those pass through as ordinary columns. Without the flag
the run refuses with `TABLE_SOURCE_COLUMN_COLLISION`.

With a single group, skip the folders and the combining runs.

## 3. Trial runs and column audit

One trial run per group, then one combined trial to audit:

```bash
npx consultchimps@0.12.0 sheets consolidate groups/row1 \
  --normalize-headers --suggest-map draft-row1.json -o trial-row1.xlsx
npx consultchimps@0.12.0 sheets consolidate groups/row3 --header-row 3 \
  --normalize-headers --suggest-map draft-row3.json -o trial-row3.xlsx
npx consultchimps@0.12.0 sheets consolidate trial-row1.xlsx trial-row3.xlsx \
  --no-source -o trial.xlsx
```

List each column of `trial.xlsx` with the suppliers that fill it. This reads the
trial output rather than transforming it, so a short script is fine; say that
you used one. In Python:

```python
import openpyxl
from collections import defaultdict
rows = openpyxl.load_workbook("trial.xlsx", read_only=True).worksheets[0].iter_rows(values_only=True)
header = next(rows); src = header.index("_source_file"); filled = defaultdict(set)
for row in rows:
    for i, value in enumerate(row):
        if value not in (None, ""): filled[header[i]].add(row[src])
for column in header: print(column, len(filled[column]), sorted(filled[column])[:6])
```

Read the list for:

- columns filled by a few suppliers each and never together: candidate synonyms
- `column_N` and `Name_2` columns: blank or duplicate headers in a source
- the same header holding different kinds of value for one supplier: an alias
  you must not apply to that supplier

## 4. Map and run

Write one mapping (see `mapping.md`) for every group, listing every column you
mean to keep as a canonical entry and the synonyms as aliases, so the group
outputs line up. Then run each group with it, and combine:

```bash
npx consultchimps@0.12.0 --json sheets consolidate groups/row1 \
  --map mapping.json -o mapped-row1.xlsx
npx consultchimps@0.12.0 --json sheets consolidate groups/row3 --header-row 3 \
  --map mapping.json -o mapped-row3.xlsx
npx consultchimps@0.12.0 --json sheets consolidate mapped-row1.xlsx mapped-row3.xlsx \
  --no-source -o consolidated/sources.xlsx
```

## 5. Reconcile and hand over

- The sum of `dataRowCount` across the inputs equals `metrics.outputRows` of the
  last run, less any annotation rows you removed. For a file with a title block,
  step 1's `dataRowCount` also counts every filled row between the title and the
  real header, and the real header itself; subtract those rows.
- `warnings` is empty, or every warning is explained.
- Tell the user: which folds you made and why, which supplier values sit under a
  folded name, any column left unfolded on purpose, and any file you ran
  separately.
