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
- **Normalized column key**: the aggressive matching form, lowercased with every
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
  contract, stating what each Excel operation does to each tracked workbook
  structure. Generated from `packages/xlsx/src/contract.ts` and checked by
  `pnpm docs:check`. Its statuses are the site's only words for a contract cell:
  "Preserved" (`preserve`), "Adjusted to stay correct" (`fix`), "Removed,
  reported as a warning" (`strip-warn`), "Refused before anything is written"
  (`refuse`), and "Needs review" (no declared cell).

## Database workspace (draft)

First draft for later review, added with the ADR 0003 foundation. Nothing here
is final; the maintainer will refine the terms and the record-editing verbs are
still to agree.

- **Workspace**: the stateful page for one project's database, opened from a
  file, edited in place, and saved back to that file.
- **Database**: the local relational store for a project, held in a single file
  that carries its own tables, schema, and identifier state.
- **Table**: a named set of columns and the records held under them.
- **Record**: one row of a table.
- **Record ID**: the human-readable, always-generated, immutable identifier for
  a record (for example `CUST-0001`) that relationships reference. Provisional
  name, agreed changeable.
- **Relationship**: a link from a column in one table to another table's Record
  ID (a foreign key).
- **Computed column** (defined for later): a column whose values are derived by
  a formula rather than entered by hand.
- **Dashboard** (defined for later): a saved arrangement of charts and figures
  drawn from the database.

## Power BI (draft)

First draft, added with ADR 0004. The verb and nouns are proposed, not final.

- **Export (Power BI)**: write each exportable table of a Power BI file's model
  to a workbook, as one or more worksheets per table. Never called "extract" or
  "convert".
- **Exportable table**: a model table the export policy admits: not hidden
  (unless hidden tables are included), within the column limit, and with at
  least one column whose values can be decoded and represented in the workbook.
- **Model**: the tables and their loaded rows carried inside a `.pbix` file. A
  template (`.pbit`) or a live-connection file carries no model.
- **Model table**: one named table of the model, with its columns and rows.
- **Hidden table**: a model table the file marks as not user-visible, such as
  the date tables Power BI generates on its own.
- **Calculated table / calculated column**: a model table or column whose rows
  were produced by a DAX expression rather than loaded from a source. Exported
  as data.
- **Manifest**: the record an export returns alongside the workbook, naming
  skipped tables and columns, worksheet splits, and counts of values rounded,
  truncated, or encoded as text, grouped by column and reason.
