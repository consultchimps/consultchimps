# @consultchimps/tabular

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
