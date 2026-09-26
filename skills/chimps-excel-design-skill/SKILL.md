---
name: chimps-excel-design-skill
description:
  Standards for building a new client-facing Excel workbook, with Excel 365
  formulas, named-range inputs, no invented figures, mandatory formatting and a
  handover. Not for editing existing files.
license: Apache-2.0
metadata:
  cli-version: "0.12.0"
  repository: consultchimps/consultchimps
---

# Design an Excel workbook deliverable

## Scope

This skill decides what a new workbook contains. The library skill you write the
file with (the built-in `xlsx` skill, or your own openpyxl or SheetJS code)
handles the mechanics, which this skill does not repeat. To change a workbook
that already exists, use `chimps-xlsx` where it is installed.

Where the library's habits differ, follow the rule on the right.

| Habit                                           | Rule here                                     |
| ----------------------------------------------- | --------------------------------------------- |
| Write a computed value into a calculation range | write the formula that produces it            |
| Point a formula at the cell holding an input    | point it at the input's named range           |
| Fill an unknown with a plausible placeholder    | leave it blank and flag it                    |
| Let the library set the document properties     | set `creator` and `lastModifiedBy` explicitly |
| Treat formatting as taste                       | the formatting rules below are mandatory      |

## Formulas

Target Excel 365. XLOOKUP, XMATCH, FILTER, UNIQUE, SORT, SORTBY, LET, LAMBDA,
SEQUENCE and TEXTSPLIT are allowed and preferred over their older equivalents.

Every function newer than Excel 2007 must be written with its `_xlfn.` prefix
when a library writes the file, or Excel opens the workbook showing `#NAME?`.
Excel hides the prefix from the user once the file is open. The per-function
table is in [references/formula-prefixes.md](references/formula-prefixes.md);
consult it rather than guessing, because two of these functions take a second
prefix segment.

One lookup family per workbook. Default to XLOOKUP. A file that already uses
INDEX/MATCH throughout stays on INDEX/MATCH. Never mix the two in one file: a
reviewer checking a formula should not have to work out which convention this
sheet follows.

## Formula-driven structure

No hardcoded figure inside a calculation range. A number a reader can see but
cannot trace is the defect this rule exists to prevent.

Every input lives on an Assumptions sheet, one input per row, with its unit, its
source and its owner beside it. Calculations reference an input by named range,
never by cell address: `Tax_Rate`, not `Assumptions!$B$7`. A named range
survives an inserted row and says what it means at the point of use.

Name ranges in `Title_Case_With_Underscores`, scoped to the workbook unless the
same name must repeat per sheet.

## Never invent data

Do not invent a figure, a score, a weight, a benchmark or a growth rate. A value
you cannot source is a blank cell, a `Needs Input` flag in the row's flag
column, and a line in the handover.

An estimate the user supplied is an input on the Assumptions sheet with its
source recorded as the user. An estimate you derived is an input too, with the
derivation recorded in its source cell.

## Formatting

These rules are mandatory in every workbook this skill produces.

- No merged cells outside a title row. Centre across selection instead: merged
  cells break sorting, filtering and range references.
- Font colour carries provenance: blue `0000FF` for a hardcoded input, black
  `000000` for a formula, green `008000` for a formula that reaches another
  sheet.
- Freeze the header row and switch on the autofilter, on every data sheet.
- One number format per column. A column is currency, or percent, or a date, and
  does not mix them.
- Column widths set so no value shows as `####`.
- No cell comments the user did not ask for. No decorative styling: no gradient
  fills, no themed banding, no icon sets used as ornament.

## Document properties

`creator` and `lastModifiedBy` must never read `openpyxl`, `SheetJS`, `xlsx` or
any other library name. Set both explicitly before saving, along with the title.
A library name in the properties tells the client who really wrote the file.

## Handover

Before delivery, give the user three things, under these three headings.

1. **Directly audited**: what you opened, recalculated or reconciled yourself,
   and against what, naming the sheets and the totals
2. **Needs a decision**: every assumption or choice that is the user's to make,
   each with the option you defaulted to and where it lives on Assumptions
3. **Gaps**: every `Needs Input` cell, every figure that could not be sourced,
   and anything the workbook models more simply than reality

An empty section is stated as empty, not omitted.

## Delivery gate

The gate is
[references/delivery-checklist.md](references/delivery-checklist.md). Run what
your environment allows, in this order, then apply the rest of the checklist
yourself. Tell the user which of these ran.

1. **ConsultChimps, when you can reach GitHub.** Each release ships the CLI as
   one file that needs only Node, no install:

   ```bash
   curl -sSLO https://github.com/consultchimps/consultchimps/releases/download/consultchimps%400.12.0/consultchimps.mjs
   node consultchimps.mjs --json sheets inspect model.xlsx
   ```

   Inspect reports each worksheet, its used range, header row, columns and
   sample values, which confirms the sheets and columns landed where you meant.
   It says nothing about formulas, formatting or properties. There is no
   `sheets check` operation yet: do not run one or say one ran.

2. **The bundled script, with or without network.** It needs Python and
   openpyxl, which ChatGPT and Claude sandboxes have, and changes nothing:

   ```bash
   python scripts/check_workbook.py model.xlsx
   ```

   It checks function prefixes, one lookup family, Assumptions cells reached by
   address, named ranges that are not defined, merged cells, font colours,
   frozen headers and autofilters, mixed number formats, dates stored as text,
   comments, and document properties, and lists every `Needs Input` cell for the
   handover. Fix each FAIL, judge each REVIEW.

3. **The checklist by hand**, for everything the two above cannot see: invented
   figures, sources and owners, hardcoded figures inside calculations, and error
   values, which appear only when Excel opens the file.

## Library pitfalls

The errors that actually reach the client are listed in
[references/library-pitfalls.md](references/library-pitfalls.md): the missing
`_xlfn.` prefix, a library name left in the document properties, merged cells,
dates written as text, and number formats lost on write. Read it before you save
the first time, not after the user reports `#NAME?`.
