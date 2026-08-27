---
"@consultchimps/tabular": minor
---

Add the column mapping engine: a declarative, versioned way to fold the same
field's different header spellings into one canonical column before tables are
consolidated.

A mapping is a version 1 JSON document of canonical columns with their aliases,
optional per-column coercions, and constant columns. `validateColumnMapping`
checks a parsed document and refuses an unusable one with
`TABLE_MAPPING_INVALID`, naming the first problem it found; reading and parsing
the file stays the caller's job, so this package still touches no filesystem.

`applyColumnMapping` (one table) and `applyColumnMappingToTables` (many) apply
it:

- Aliases match by normalized column key, so one alias entry catches every case,
  spacing, and punctuation variant of that spelling, independently of the
  `normalizeHeaders` union option. Canonical names are written verbatim.
- Unmapped columns pass through under their own names and come back in
  `unmappedColumns`, so an operation can warn about them - loud but lossless.
- Two columns of one table folding into one canonical column is refused with
  `TABLE_MAPPING_COLUMN_COLLISION` naming the file, sheet, both columns, and the
  canonical column, because combining them would silently drop values.
- A declared date format is parsed into an ISO 8601 date string, and declared
  thousands and decimal separators into a number. Separator placement is checked
  before any separator is removed, so a malformed grouping such as "1,23.4" is
  refused rather than silently read as a different amount. A value that is not
  what its coercion declares is refused with `TABLE_MAPPING_COERCION_FAILED`
  naming the row and column; blank cells stay blank.
- Constant columns are appended after the mapped and unmapped columns, and a
  constant column that collides with an existing one is refused with
  `TABLE_MAPPING_CONSTANT_COLLISION`.

`suggestColumnMapping` drafts a mapping from the inputs' header lists by
grouping columns whose normalized keys already match, proposing the first
spelling seen as the canonical name and returning the evidence behind each
group. It uses no similarity scoring and no sample values, so it has no false
positives; synonyms that share no spelling stay a manual mapping entry, and
nothing is applied without review.

Column matching is now independent of the machine's locale. `columnKey` and
`normalizedColumnKey` case-fold with `toLowerCase` instead of the locale-aware
`toLocaleLowerCase`, which read the host's default locale: under a Turkish or
Azeri locale, "ID" folded to "ıd", so consolidating the same workbooks could
align headers on one machine and split them into separate columns on another.
Matching now behaves the same everywhere, which changes how headers containing a
dotted or dotless I are matched on hosts set to those locales.
