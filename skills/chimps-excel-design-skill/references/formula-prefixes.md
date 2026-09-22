# The `_xlfn.` prefix

A function added after Excel 2007 has no entry in the file format's built-in
function table. Excel stores it by name with a `_xlfn.` prefix and strips the
prefix when it renders the formula. Excel adds the prefix for you; a library
does not. A library that writes `=XLOOKUP(...)` produces a workbook that opens
showing `#NAME?` in every cell that used it, and the user cannot repair it by
retyping, because the workbook is what is wrong.

The rule cuts both ways. Prefixing a function that Excel 2007 already had, such
as `SUM`, `IF`, `INDEX`, `MATCH`, `VLOOKUP`, `SUMIFS` or `IFERROR`, also gives
`#NAME?`. Prefix the new ones only.

## Write it exactly like this

| Function   | Written as           | Notes                             |
| ---------- | -------------------- | --------------------------------- |
| XLOOKUP    | `_xlfn.XLOOKUP`      | the default lookup here           |
| XMATCH     | `_xlfn.XMATCH`       |                                   |
| FILTER     | `_xlfn._xlws.FILTER` | second prefix segment, not a typo |
| SORT       | `_xlfn._xlws.SORT`   | second prefix segment, not a typo |
| SORTBY     | `_xlfn.SORTBY`       | one segment, unlike SORT          |
| UNIQUE     | `_xlfn.UNIQUE`       |                                   |
| SEQUENCE   | `_xlfn.SEQUENCE`     |                                   |
| RANDARRAY  | `_xlfn.RANDARRAY`    |                                   |
| LET        | `_xlfn.LET`          | names inside take no prefix       |
| LAMBDA     | `_xlfn.LAMBDA`       | parameters take no prefix         |
| ISOMITTED  | `_xlfn.ISOMITTED`    |                                   |
| TEXTSPLIT  | `_xlfn.TEXTSPLIT`    |                                   |
| TEXTBEFORE | `_xlfn.TEXTBEFORE`   |                                   |
| TEXTAFTER  | `_xlfn.TEXTAFTER`    |                                   |
| VSTACK     | `_xlfn.VSTACK`       |                                   |
| HSTACK     | `_xlfn.HSTACK`       |                                   |
| TOCOL      | `_xlfn.TOCOL`        |                                   |
| TOROW      | `_xlfn.TOROW`        |                                   |
| CHOOSECOLS | `_xlfn.CHOOSECOLS`   |                                   |
| CHOOSEROWS | `_xlfn.CHOOSEROWS`   |                                   |
| BYROW      | `_xlfn.BYROW`        |                                   |
| BYCOL      | `_xlfn.BYCOL`        |                                   |
| MAP        | `_xlfn.MAP`          |                                   |
| REDUCE     | `_xlfn.REDUCE`       |                                   |
| SCAN       | `_xlfn.SCAN`         |                                   |
| MAKEARRAY  | `_xlfn.MAKEARRAY`    |                                   |
| IFS        | `_xlfn.IFS`          | Excel 2016, still prefixed        |
| SWITCH     | `_xlfn.SWITCH`       | Excel 2016, still prefixed        |
| MAXIFS     | `_xlfn.MAXIFS`       | Excel 2016, still prefixed        |
| MINIFS     | `_xlfn.MINIFS`       | Excel 2016, still prefixed        |
| CONCAT     | `_xlfn.CONCAT`       | Excel 2016, still prefixed        |
| TEXTJOIN   | `_xlfn.TEXTJOIN`     | Excel 2016, still prefixed        |

Two operators are stored the same way. A spilled-range reference, typed as
`E2#`, is written `_xlfn.ANCHORARRAY(E2)`. An implicit intersection, typed as
`@A:A`, is written `_xlfn.SINGLE(A:A)`.

A function not listed here that postdates Excel 2007 almost certainly needs the
prefix too. Confirm before shipping: write one cell using it, open the result in
Excel, and look for `#NAME?`.

## Checking your own output

The prefix is a string, so it is checkable without Excel. Search the formulas
you wrote for the bare names in the left column: any hit outside a prefixed
occurrence is a cell that will fail. Do this before delivery, because the error
does not appear until someone opens the file.
