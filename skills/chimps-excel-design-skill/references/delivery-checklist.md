# Delivery checklist

Apply this before handing the workbook over. There is no ConsultChimps operation
that checks an authored workbook: a `sheets check` operation is planned and does
not exist yet, so this list is the gate.

Two tools cover part of it. `scripts/check_workbook.py`, a heuristic lint for
the common library mistakes rather than a proof, checks the items marked
(script) below from the file alone, and the part of each (script, partly) item
that the file shows: Assumptions cells reached by address, formats of filled
cells, and comments. ConsultChimps `sheets inspect`, run from the release file
on GitHub, confirms the sheets, headers and columns are the ones you meant.
Everything unmarked is read by you.

## Formulas

- [ ] (script, partly) Every function newer than Excel 2007 carries its `_xlfn.`
      prefix, FILTER and SORT carry `_xlfn._xlws.`, and no older function
      carries a prefix
- [ ] (script) One lookup family across the whole workbook
- [ ] No hardcoded figure inside a calculation range
- [ ] (script, partly) Every input reference uses a named range, not a cell
      address
- [ ] (script) Every named range used in a formula is registered in the
      workbook's defined names
- [ ] No error value visible anywhere: open the file and look

## Data honesty

- [ ] No figure, score, weight or rate that you invented
- [ ] Every unsourced value is a blank cell plus a `Needs Input` flag
- [ ] Every `Needs Input` cell appears in the handover
- [ ] Every input on Assumptions states its unit, source and owner

## Formatting

- [ ] (script) No merged cells outside a title row
- [ ] (script) Input cells blue, formula cells black, cross-sheet formulas green
- [ ] (script) Header row frozen and autofilter on, on every data sheet
- [ ] (script, partly) One number format per column, applied to blank cells too
- [ ] No column showing `####`
- [ ] (script, partly) No cell comments the user did not ask for, and no
      decorative styling

## File

- [ ] (script) `creator` and `lastModifiedBy` set, neither reading as a library
      name
- [ ] (script) Title set
- [ ] Opened in Excel once, after the final save

## Handover

- [ ] Directly audited: what you checked yourself, against what
- [ ] Needs a decision: each open assumption, its default, its location
- [ ] Gaps: every Needs Input cell, unsourced figure and simplification
- [ ] Any empty section stated as empty rather than dropped
