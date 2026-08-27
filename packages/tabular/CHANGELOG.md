# @consultchimps/tabular

## 0.4.0

### Minor Changes

- 7f309f6: Add the column mapping engine: a declarative, versioned way to fold
  the same field's different header spellings into one canonical column before
  tables are consolidated.

  A mapping is a version 1 JSON document of canonical columns with their
  aliases, optional per-column coercions, and constant columns.
  `validateColumnMapping` checks a parsed document and refuses an unusable one
  with `TABLE_MAPPING_INVALID`, naming the first problem it found; reading and
  parsing the file stays the caller's job, so this package still touches no
  filesystem.

  `applyColumnMapping` (one table) and `applyColumnMappingToTables` (many) apply
  it:

  - Aliases match by normalized column key, so one alias entry catches every
    case, spacing, and punctuation variant of that spelling, independently of
    the `normalizeHeaders` union option. Canonical names are written verbatim.
  - Unmapped columns pass through under their own names and come back in
    `unmappedColumns`, so an operation can warn about them - loud but lossless.
  - Two columns of one table folding into one canonical column is refused with
    `TABLE_MAPPING_COLUMN_COLLISION` naming the file, sheet, both columns, and
    the canonical column, because combining them would silently drop values.
  - A declared date format is parsed into an ISO 8601 date string, and declared
    thousands and decimal separators into a number. Separator placement is
    checked before any separator is removed, so a malformed grouping such as
    "1,23.4" is refused rather than silently read as a different amount. A value
    that is not what its coercion declares is refused with
    `TABLE_MAPPING_COERCION_FAILED` naming the row and column; blank cells stay
    blank.
  - Constant columns are appended after the mapped and unmapped columns, and a
    constant column that collides with an existing one is refused with
    `TABLE_MAPPING_CONSTANT_COLLISION`.

  `suggestColumnMapping` drafts a mapping from the inputs' header lists by
  grouping columns whose normalized keys already match, proposing the first
  spelling seen as the canonical name and returning the evidence behind each
  group. It uses no similarity scoring and no sample values — only deterministic
  normalization equivalence, which can still group headers that differ only in
  punctuation, so every draft is meant for review; synonyms that share no
  spelling stay a manual mapping entry, and nothing is applied without review.

  Column matching is now independent of the machine's locale. `columnKey` and
  `normalizedColumnKey` case-fold with `toLowerCase` instead of the locale-aware
  `toLocaleLowerCase`, which read the host's default locale: under a Turkish or
  Azeri locale, "ID" folded to "ıd", so consolidating the same workbooks could
  align headers on one machine and split them into separate columns on another.
  Matching now behaves the same everywhere, which changes how headers containing
  a dotted or dotless I are matched on hosts set to those locales.

### Patch Changes

- 32973f7: Declare support for Node.js 22 and later (`engines.node: ">=22.0.0"`
  instead of `24.x`), so the toolkit installs and runs in environments that ship
  the previous LTS line. CI now validates the runtime on Node 22.16, the latest
  22, and 26 alongside the full Node 24 verification. The CLI additionally ships
  a standalone `consultchimps.mjs` bundle on each GitHub release: one file that
  runs with `node consultchimps.mjs` and needs no npm access at all.
- Updated dependencies [32973f7]
  - @consultchimps/core@0.5.1

## 0.3.0

### Minor Changes

- 98c77a7: Consolidation can now match columns whose headers differ only in
  case, spacing, or punctuation. Different systems often export the same schema
  with different header conventions - "Failed Checks" in one file,
  "Failed_Checks" in another, "Reviewer: Lead Contact" versus
  "Reviewer_Lead_Contact" - and until now each spelling became its own
  mostly-empty output column.

  Opt in with `--normalize-headers` on `consultchimps sheets consolidate`, or
  `normalizeHeaders: true` on `consolidateWorkbooks` and `unionTables`. Matching
  ignores case and treats any run of spaces or punctuation as one separator; the
  first spelling seen names the output column. The default behaviour is
  unchanged. The tabular package also exports the new `normalizedColumnKey`
  helper behind this matching.

## 0.2.4

### Patch Changes

- Updated dependencies [1f759eb]
  - @consultchimps/core@0.5.0

## 0.2.3

### Patch Changes

- Updated dependencies [6564e24]
  - @consultchimps/core@0.4.0

## 0.2.2

### Patch Changes

- Updated dependencies [5c08c75]
  - @consultchimps/core@0.3.0

## 0.2.1

### Patch Changes

- Updated dependencies [c78b35e]
  - @consultchimps/core@0.2.0

## 0.2.0

### Minor Changes

- 7958e80: Add reusable table grouping and an Excel split-by-column API and CLI
  command with deterministic filenames, blank-row controls, and safe output
  preflight.
