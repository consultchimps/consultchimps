# What breaks a library-written workbook

This is guidance, not a procedure. The library your chat has is the one to use;
these are the five failures that reach the client, whichever library that is.
openpyxl is named where the behaviour is specific to it.

## 1. The missing `_xlfn.` prefix

A modern function written without its prefix opens as `#NAME?`, and every cell
that depends on it shows `#NAME?` too, so one wrong string can read as a broken
workbook. See [formula-prefixes.md](formula-prefixes.md). This is the single
most common way a workbook that looked right in code arrives wrong.

## 2. A library name in the document properties

openpyxl stamps its own name into the workbook's creator, and every save
rewrites `lastModifiedBy`. The client opens the file properties and reads
`openpyxl`. Set creator, lastModifiedBy and title explicitly, as the last step
before saving, and set them again after any later save.

## 3. Merged cells

A merged range stores its value in the top-left cell only; the rest read as
empty. Sorting, filtering, XLOOKUP over the range and any formula that walks the
column then see blanks, and Excel itself refuses some operations over merged
cells. Merge only a title row. For a heading that should sit visually across
columns, use centre across selection, which is an alignment setting rather than
a structural change.

## 4. Dates written as text

A date written as a string is left-aligned text that looks like a date. It sorts
alphabetically, fails a date filter, and never matches a real date in a lookup,
silently returning the not-found branch. Write a real date value and apply a
date number format to the cell. Times carry the same trap, plus the reverse one:
a value read from a source as a serial number needs a date format, or it shows
as 45993.

## 5. Number formats lost on write

Formats belong to cells, not to columns or to data. Writing a value into a cell
that was never formatted leaves it General, and a row appended below the last
formatted row inherits nothing. Apply the column's format to every cell you
write, including the blank ones a Needs Input flag leaves behind, so the column
stays one format. Round-tripping an existing workbook through a library loses
what the library does not model: charts, pivot tables, slicers, images and some
conditional formatting can disappear on save, which is one reason this skill
authors new files rather than rewriting supplied ones.

## Also worth knowing

- A library writes formulas as text and computes nothing, so the file carries no
  cached results. Excel calculates on open; other readers may show blank or
  zero. Never read your own formula's result back from the file you just wrote,
  and never quote such a number to the user as a computed total.
- A named range must be registered in the workbook's defined names. Writing the
  name into a formula is not enough; an unregistered name is `#NAME?`.
- Freeze panes and the autofilter are per-worksheet settings. Set both on every
  data sheet; neither is inherited from another sheet.
- A dynamic array formula is written once, in the top-left cell of where it will
  spill. Writing it into every cell of the intended range produces a spill
  collision in each one.
