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

`--header-row` applies to every file in a run. Files whose real headers sit on
row 1 go in one run; files with a two row title block go in a run with
`--header-row 3`, and so on. Pass the same `--map` to every group run so the
columns line up.

Combine the group outputs in a last run with `--no-source`. Each group output
already carries `_source_file`, `_source_sheet` and `_source_row` naming the
original workbooks, and those pass through as ordinary columns. Without the flag
the run refuses with `TABLE_SOURCE_COLUMN_COLLISION`.

## 3. Trial run and column audit

```bash
npx consultchimps@0.12.0 sheets consolidate templates/sources \
  --normalize-headers --suggest-map draft.json -o trial.xlsx
```

List each output column with the suppliers that fill it. This reads the trial
output rather than transforming it, so a short script is fine; say that you used
one. In Python:

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

Write the mapping (see `mapping.md`), listing every column you mean to keep as a
canonical entry and the synonyms as aliases. Then:

```bash
npx consultchimps@0.12.0 --json sheets consolidate templates/sources \
  --map mapping.json -o "consolidated/sources.xlsx"
```

## 5. Reconcile and hand over

- The sum of `dataRowCount` across the inputs equals `metrics.outputRows`, less
  any annotation rows you removed.
- `warnings` is empty, or every warning is explained.
- Tell the user: which folds you made and why, which supplier values sit under a
  folded name, any column left unfolded on purpose, and any file you ran
  separately.
