---
"@consultchimps/xlsx": minor
"@consultchimps/messages": minor
"consultchimps": minor
---

A workbook split to CSV with no table, range or worksheet named now splits every
worksheet that carries the column, instead of refusing a workbook with several.
Each worksheet gives one CSV file per value, named
`<prefix>-<value> - <sheet>.csv` when more than one worksheet is split, with the
value and sheet name each made safe and a `-2` suffix for any name that would
repeat. A worksheet without the column is left out with a warning, the plan
lists every file, and keeping the workbook stays refused with
`XLSX_SPLIT_CSV_PRESERVE`. A workbook whose hidden worksheet also carries the
column now gives that worksheet's files too, so its files are named for their
worksheets.
