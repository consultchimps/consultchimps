# Delivery checklist

Apply this before handing the workbook over. There is no ConsultChimps operation
that checks an authored workbook: a `sheets check` operation is planned and does
not exist yet, so this list is the gate.

`consultchimps sheets inspect <file>` covers one part of it. It reports each
worksheet, its used range, the header row, the columns and up to five sample
values per column, so it confirms that the sheets, headers and columns are the
ones you meant to produce. It reports nothing about formulas, formatting or
document properties, so the rest of this list is read by you.

## Formulas

- [ ] Every function newer than Excel 2007 carries its `_xlfn.` prefix, FILTER
      and SORT carry `_xlfn._xlws.`, and no older function carries a prefix
- [ ] One lookup family across the whole workbook
- [ ] No hardcoded figure inside a calculation range
- [ ] Every input reference uses a named range, not a cell address
- [ ] Every named range used in a formula is registered in the workbook's
      defined names
- [ ] No error value visible anywhere: open the file and look

## Data honesty

- [ ] No figure, score, weight or rate that you invented
- [ ] Every unsourced value is a blank cell plus a `Needs Input` flag
- [ ] Every `Needs Input` cell appears in the handover
- [ ] Every input on Assumptions states its unit, source and owner

## Formatting

- [ ] No merged cells outside a title row
- [ ] Input cells blue, formula cells black, cross-sheet formulas green
- [ ] Header row frozen and autofilter on, on every data sheet
- [ ] One number format per column, applied to blank cells too
- [ ] No column showing `####`
- [ ] No cell comments the user did not ask for, and no decorative styling

## File

- [ ] `creator` and `lastModifiedBy` set, neither reading as a library name
- [ ] Title set
- [ ] Opened in Excel once, after the final save

## Handover

- [ ] Directly audited: what you checked yourself, against what
- [ ] Needs a decision: each open assumption, its default, its location
- [ ] Gaps: every Needs Input cell, unsourced figure and simplification
- [ ] Any empty section stated as empty rather than dropped
