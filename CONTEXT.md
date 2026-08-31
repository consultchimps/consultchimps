# Glossary

The canonical vocabulary for ConsultChimps operations. Code, docs, CLI help, and
issues use these terms with exactly these meanings.

- **Consolidate**: stack rows from many worksheets into one table. Never called
  "merge".
- **Merge (workbooks)**: copy worksheets into one workbook as separate tabs.
  Never stacks rows.
- **Split (spreadsheets)**: produce one workbook per distinct value of a column.
- **Split (PDF)**: produce one file per page of a document.
- **Inspect**: describe an input's structure without producing files. The
  toolkit's single verb for this, used by PowerPoint template inspection and
  workbook inspection alike ("describe" appears only in library function names
  such as `describeWorkbook`).
- **Column key**: the case-folded, trimmed form of a header used for matching
  (`columnKey`).
- **Normalized column key**: the aggressive matching form: lowercased, every
  separator run collapsed to one underscore (`normalizedColumnKey`). Matching
  only; never shown as output.
- **Canonical column**: the output column name a mapping declares. Written
  verbatim to the output, never normalized.
- **Alias**: a source header spelling that a mapping folds into a canonical
  column. Aliases match by normalized column key.
- **Column mapping**: the declarative, versioned JSON document of canonical
  columns, their aliases, optional coercions, and constant columns, applied
  during consolidation.
- **Coercion**: a deterministic per-column value conversion declared in a
  mapping (dates from a declared format, number parsing).
- **Unmapped column**: an input column no mapping entry matches. Passes through
  under its own name, with a warning.
- **Assist / suggestion**: a drafted mapping produced from
  normalization-equivalence groups for the user to review. Never applied
  silently.
- **Surface**: one of the three ways an operation ships (CLI, library, browser),
  each with its own status in the tool registry (ADR 0001).
- **Source columns**: the provenance columns consolidation appends
  (`_source_file`, `_source_sheet`, `_source_row`).
- **Preservation matrix**: the published projection of the xlsx conformance
  contract: what each Excel operation does to each tracked workbook structure.
  Generated from `packages/xlsx/src/contract.ts` and checked by
  `pnpm docs:check`. Its statuses are the site's only words for a contract cell:
  "Preserved" (`preserve`), "Adjusted to stay correct" (`fix`), "Removed,
  reported as a warning" (`strip-warn`), "Refused before anything is written"
  (`refuse`), and "Needs review" (no declared cell).
