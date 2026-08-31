# Workbook inspection and column mapping

Consolidation matches columns by header name, so inputs from different systems
that spell the same field differently union into duplicate columns (a real
fourteen-sheet run produced a sixty-nine-column union). We decided to ship two
cooperating features: a workbook **inspection** operation (`sheets.inspect`)
that describes a workbook's structure on all three surfaces, and a declarative
**column mapping** applied during consolidation, with inspection's output
feeding the mapping review UI.

## Decisions

- **Inspection is a first-class operation** on the library (`describeWorkbook` /
  `describeWorkbookBytes`), the CLI (`sheets inspect`), and the browser
  (`/tools/excel-inspect`), mirroring `pptx.inspect-template` as the
  create-nothing precedent. Its result carries structure (sheets with visibility
  and dimensions, header previews, Excel Tables, named ranges) plus **bounded**
  per-column sample values, because samples are what mapping review and the tool
  pickers need; unbounded values or inferred types are deliberately excluded.
  The bytes surface also gains the Excel Tables and named-range readers it was
  missing, making the previously dangling type re-exports truthful.
- **The mapping engine lives in `@consultchimps/tabular`**, not the xlsx
  package: header mapping is format-independent, and this placement keeps the
  door open to mapping tables extracted from other document types without a
  migration.
- **Mapping matches via `normalizedColumnKey`, always.** One alias entry catches
  every spelling variant; canonical names are written verbatim.
  `normalizeHeaders` continues to govern only how unmapped columns union. We
  rejected exact matching (a large mapping would have to enumerate every
  variant) and coupling mapping to the `normalizeHeaders` flag.
- **The v1 mapping file is versioned JSON, global-scope**: canonical column →
  aliases, optional per-column coercions (dates from a declared format,
  numbers), and constant columns. YAML and per-source override sections are
  deferred, not rejected: they are additive later. JSON was chosen to add zero
  dependencies. Validation rejects a mapping whose aliases collide on normalized
  keys, within one canonical column or across two, before any data is read,
  since matching is normalized and such a document is ambiguous by construction.
- **Unmapped columns pass through with a warning** listing them: loud but
  lossless. **Two same-sheet columns mapping to one canonical is a refusal**
  with a stable error naming the sheet, columns, and canonical: merging same-row
  values silently would be data loss, and the house rule is to fail before
  writing on ambiguity. Coalescing non-empty values was considered and deferred
  as value-level merging beyond v1.
- **Assist mode suggests only normalization-equivalence groups** (columns whose
  normalized keys already match), drafted by `--suggest-map` and by the
  consolidate tool page, and **never applied silently**. String similarity and
  sample-value overlap were considered and rejected for v1: the mission requires
  deterministic, explainable behavior, and equivalence grouping proposes no
  matches beyond spelling variants of one name. Normalization can still conflate
  punctuation-distinct headers ("A+B" and "A-B" both normalize to `a_b`), one
  more reason every suggestion is reviewed, never applied silently. Synonym
  headers ("Timestamp" vs "Run Time") remain a manual mapping entry.

## Consequences

- The inspect page's picker component is shared: split and consolidate embed it
  after the standalone page ships.
- The mapping JSON is a public file format; changes to it are versioned public
  API changes.
- `sheets.inspect` joins the tool registry with all three surfaces and is bound
  by the feature-completion checklist per surface.
