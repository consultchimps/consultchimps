# In-browser tool smoke tests

Playwright drives the statically exported site, not the dev server, so these
tests exercise the same bundles GitHub Pages serves, including the operation Web
Worker, the engines it imports on demand, and the blob download path.

## Running the suite

The suite serves `apps/docs/out`, so the export has to exist first, and it must
be built **without** `NEXT_PUBLIC_BASE_PATH` so the routes are served from `/`:

```bash
pnpm --filter @consultchimps/docs... build
pnpm --filter @consultchimps/docs exec playwright install chromium
pnpm --filter @consultchimps/docs e2e
```

`pnpm --filter @consultchimps/docs e2e` fails with a pointer to the build
command when `out/` is missing.

## What is covered

- `pdf-split.spec.ts`: splitting a two-page PDF into zero-padded page files,
  downloading one of them, and refusing a file that is not a PDF.
- `pdf-merge.spec.ts`: merging two single-page PDFs into `combined.pdf` and
  downloading the result.
- `excel-split.spec.ts`: detecting a workbook's column headers, splitting on one
  of them into a workbook per distinct value, downloading one of them, reading a
  downloaded workbook back to confirm it keeps every source worksheet and
  removes only the other values' rows, reporting a column the workbook does not
  have, inspecting the chosen workbook from the page, and refusing a file that
  is not a workbook.
- `excel-merge.spec.ts`: merging two workbooks into one, reordering and removing
  sources, and downloading the result.
- `excel-consolidate.spec.ts`: stacking two workbooks whose headers drifted
  apart into one table, downloading the result, checking that the "Normalize
  headers" and "Add source columns" checkboxes change the columns the finished
  workbook holds, inspecting one of the added workbooks, applying a column
  mapping end to end and reading the consolidated headers back, refusing an
  ambiguous mapping on selection with its stable error reference, and drafting a
  mapping: the proposed groups, the reviewed draft as valid version 1 JSON, and
  the round trip that applies a draft only after it is added back.
- `excel-inspect.spec.ts`: describing a workbook's worksheets, hidden tabs,
  header rows, Excel Table, named range, and sample column values, turning the
  hidden-worksheet option off to get the description an operation would see,
  inspecting a macro-enabled workbook, refusing a file that is not a workbook,
  explaining a workbook the picker accepts but the operation cannot read (the
  stable error reference included), shortening a 120-column worksheet until the
  toggle asks for the rest, and replacing a workbook mid-inspection so the
  withdrawn report cannot describe the file that was replaced.
- `pptx.spec.ts`: populating a template slide from workbook records into one
  deck, naming the output, downloading it, reporting a placeholder no column
  feeds, refusing a template that is not a presentation, and inspecting a
  template's placeholders with their occurrence counts on the chosen slide.
- `workspace.spec.ts`: reaching `/workspace` from the header, starting an empty
  workspace, saving it to a `.sqlite` file, reopening exactly those bytes, and
  reporting a file that is not a readable database or not a workspace type.
- `workspace-import.spec.ts`: importing one worksheet of a two-sheet workbook
  and a `.csv` into an open workspace, checking the table listing that follows
  (row counts, Record ID prefixes, and the inferred column types), keeping both
  tables across a save and a reopen, refusing a table name the workspace already
  reads as the same name, and reporting a file with nothing to import. It also
  covers the shell's unsaved-changes guard, because import is the first command
  that can leave a workspace holding work no file has: New and Open ask before
  replacing an imported workspace and leave it untouched until the loss is
  confirmed, they ask nothing once it has been saved, and a link out of the page
  and the Back button are both held the same way. The guard covers an import
  still in flight as well as unsaved changes, so a link click during one is held
  even from a clean workspace. A Back press that finds a spare entry with
  something at stake spends it and arms another, so a second press is caught
  too; one that finds a spare with nothing left to guard retires it and reaches
  the page before this one, rather than being spent on nothing. An import that
  fails under a question raised by a held link dismisses that question, because
  the reason it was raised about did not happen. Further tests hold the worker's
  import commands open to check that New, Open, and Save are disabled while a
  file is being read and while an import runs, and refuse a worksheet whose
  formulas the workbook carries no calculated value for.
- `workspace-grid.spec.ts`: showing a table's records in the grid, switching
  between tables, editing text, number, and foreign-key cells and finding those
  edits in the saved and reopened file, and having an impossible value and an
  emptied non-nullable column both refused, reverted, and explained. The Record
  ID column offers no editor by click or by tab, and a column name with a dot in
  it is addressed as the column it is rather than as a path into nested data.
  The rest covers where the grid meets the shell: an edit marks the workspace
  unsaved so New asks first, a saved edit makes it ask nothing, editing is
  locked while an import is in flight, and a table imported after the grid was
  already up appears in its switcher and is editable there. A Back press with a
  cell still open for editing is held: the confirmation names the draft, the
  editor stays open behind it because a standing question does not lock the
  grid, keeping the workspace keeps what was typed, and discarding leaves for
  the page before this one. Its workspace fixture is built in Node with
  `@consultchimps/db` and opened through the page's own file input, so the spec
  exercises the grid without waiting on import.
- `workspace-grid-gestures.spec.ts`: the grid's spreadsheet gestures. A range
  copied as the tab and CRLF text Excel reads, a block pasted back from a
  Windows copy, a paste anchored on the last record refused whole because
  records are not added by a paste, a number series and a date series and a
  text-with-a-number series filled from the corner of a selection, a sideways
  fill where one column takes the value and the whole-number column cannot (the
  others accepted, that one left as it was and explained), a gesture naming a
  workspace the worker does not hold refused whole with nothing written, and a
  single click selecting where a double click edits. The clipboard is driven by
  dispatching the browser's own copy and paste events with a `DataTransfer` on
  them, rather than through the system clipboard, which needs permissions the
  export never asks for.
- `tools-navigation.spec.ts`: the `/tools` index, the sub-bar tabs, the
  tool-named "Try ... online" button each guide gains from the tool registry,
  and the single button a guide shared by two operations offers.

Every downloaded PDF is checked for the `%PDF-` header and every downloaded
workbook or presentation for the `PK` ZIP header (both `.xlsx` and `.pptx` are
ZIP packages), so a tool that "finishes" while producing empty or corrupt bytes
fails the suite. `readWorkbookDownload` goes further and opens a downloaded
workbook with jszip, resolving each worksheet through the workbook's own
relationships and its shared-string table, so a test can assert which
worksheets, which headers, and which rows reached the user. `readTextDownload`
does the same for the documents the pages build themselves, such as the
consolidate page's reviewed mapping draft.

## Selectors

Address the tool pages through the `data-testid` attributes the shared tool
shell renders, not through heading text. The stable identifiers are:

| Identifier                            | Element                                  |
| ------------------------------------- | ---------------------------------------- |
| `file-picker` / `file-input`          | the drop zone and its file input         |
| `source-summary`                      | the chosen single input's name and size  |
| `source-list` / `source-item`         | the ordered list of merge inputs         |
| `preview-section` / `preview-error`   | the plan preview and its failure text    |
| `planned-outputs`                     | the planned output names                 |
| `run-button` / `cancel-button`        | the run controls                         |
| `progress-report`                     | the progress bar and its labels          |
| `results-section`                     | the Results region                       |
| `artifact-list` / `artifact-item`     | the produced outputs                     |
| `artifact-name` / `artifact-download` | one output's name and Download button    |
| `archive-download`                    | "Download all (.zip)", multi-output only |
| `result-message` / `failure-message`  | the outcome text, by outcome             |

The Excel split page adds `column-select`, `column-input`, and one identifier
per advanced control (`prefix-input`, `sheet-input`, `table-input`,
`range-input`, `header-row-input`, `include-blank-checkbox`,
`include-hidden-checkbox`, `preserve-workbook-checkbox`, `strict-checkbox`,
`values-checkbox`).

The Excel merge page adds `output-name-input` and `values-checkbox`; the Excel
consolidate page adds `output-name-input`, `normalize-headers-checkbox`,
`source-columns-checkbox`, and `include-hidden-checkbox`. Both arrange their
inputs through the "Move X earlier", "Move X later", and "Remove X" buttons on
each `source-item`.

The consolidate page takes a second kind of file, so it renders two
`file-input`s: address them through `sectionFileInput(page, "source-section")`
and `sectionFileInput(page, "mapping-section")` rather than the bare helper. The
mapping section reports into `mapping-reading`, `mapping-summary`,
`mapping-columns`, `mapping-rejected`, `mapping-error` (the engine's refusal,
stable error reference included), and `mapping-remove`. Below them,
`suggest-button` drafts a mapping into `suggestion-list`, one `suggestion-group`
per proposal carrying `suggestion-spellings`, `suggestion-evidence`, and the
editable `suggestion-canonical`; `suggestion-download` hands back the reviewed
document and reports a review the engine refuses in `suggestion-error`,
`suggest-error` holds a failed drafting, and `suggestion-empty` stands in when
no headers need folding together.

Both the split and consolidate pages fold the shared `WorkbookInspector` into
`inspector-disclosure`, whose `summary` opens it; the consolidate page chooses
which workbook to describe through `inspect-select`. The report is mounted only
while the disclosure is open, so `inspection-section` and everything under it
exist only after that click.

The Excel inspect page has a single `source-section` (with `source-summary`,
`source-reading`, `source-rejected`, and `include-hidden-checkbox`) and reports
into `inspection-section`, the shared `WorkbookInspector`. That section renders
the metric tiles `inspection-worksheets`, `inspection-data-rows`,
`inspection-excel-tables`, and `inspection-named-ranges`, then
`hidden-worksheets-callout` when the description covers a hidden worksheet, then
`worksheet-list` with one `worksheet-item` per worksheet, each carrying a
`worksheet-name`, a `worksheet-visibility` badge when it is not visible, a
`worksheet-summary`, and a `column-list` of `column-item` entries holding a
`column-header` and its `sample-value` chips. A worksheet wider than the preview
limit renders only the first hundred columns, with `column-preview-note` naming
both counts and `column-toggle` revealing the rest and putting them away again.
`excel-table-list` and `named-range-list` follow, with `no-excel-tables` and
`no-named-ranges` in their place when the workbook declares none, and
`inspection-warnings` holds one `inspection-warning` per condition the operation
reported, or `inspection-error` when the workbook cannot be read. Like the
PowerPoint inspect page it has no Run button and no Results panel: the operation
creates nothing, so the report is the whole answer.

The PowerPoint populate page takes two files, so it wraps each picker in its own
section: `template-section` (with `template-summary`) and `records-section`
(with `records-summary` and the advanced controls `worksheet-input`,
`header-row-input`, `template-slide-input`, `output-name-input`). Because both
sections render a `file-input`, always scope the input to its section on that
page rather than using the bare `file-input` helper. The PowerPoint inspect page
has a single `source-section` (with `source-summary` and `template-slide-input`)
and reports into `inspection-section`, which renders `placeholder-list` with one
`placeholder-item` per placeholder, each carrying a `placeholder-name` and its
occurrence count, plus `inspection-warnings` holding one `inspection-warning`
per condition that would make a populate refuse the template, or
`inspection-error` when the template cannot be read. That page has no Run
button: choosing a template inspects it after the usual preview debounce.

Both PowerPoint pages reject a slide or row number that is not a whole number
counted from 1 rather than falling back to a default. The offending field
renders `<field>-error` (`template-slide-input-error`, `header-row-input-error`)
and the task is withdrawn: the populate page also lists the messages in
`preview-invalid-options` and disables `run-button`, and the inspect page shows
`inspection-invalid-slide` and clears the report.

Both pages also clear the chosen file and render `template-rejected` (or
`records-rejected`) when a picker is handed something it cannot read, rather
than keeping the previous document. Choosing a file clears the previous
selection immediately and shows `template-reading` / `records-reading` /
`source-reading` until the read finishes, so Run is never enabled against a
document that has already been replaced. And because a changed option applies to
Run at once, a preview or report is shown only while it still matches the page:
changing an input replaces it with `preview-pending` or `inspection-pending`
until the recomputed answer arrives. Both transient states last at least the 250
ms preview debounce, so they are safe to assert. A finished deck is withdrawn on
the same rule: changing any option removes `results-section` entirely, because a
worksheet change can produce a deck with the identical filename from different
rows.

The workspace page is not an operation, so it renders none of the shell
identifiers above. It starts with `workspace-actions` (`workspace-new`,
`workspace-open`, and the fallback `file-input` for browsers without the File
System Access API), then either `workspace-empty` or `workspace-summary`. The
summary holds `workspace-file-name`, `workspace-table-count`, `workspace-save`,
and `workspace-save-as`, and lists what the workspace holds: one
`workspace-table` per table carrying `workspace-table-name`,
`workspace-table-rows`, `workspace-table-prefix`, and `workspace-table-columns`,
with `workspace-tables-empty` in their place while there are none.
`workspace-unsaved` is the badge shown while the workspace holds changes no file
has, and `workspace-notice` and `workspace-error` report the outcome of the last
command.

While that badge is showing, New, Open, a link out of the page, and the Back
button all render `workspace-confirm` instead of leaving, with
`workspace-confirm-discard` and `workspace-confirm-cancel` as its two answers.
It is an inline section rather than a `window.confirm`, so it is asserted like
any other part of the page. A source the import cannot read renders
`workspace-import-blocked` in its row and refuses the tick.

The record grid renders `workspace-grid-section`, holding
`workspace-table-select` (the table switcher, disabled while the shell is busy),
`workspace-grid-loading` while a table is being read, `workspace-grid-empty`
when the workspace has no tables, the grid itself under `workspace-grid`, and
`workspace-grid-reference-note` when a foreign-key picker lists only the first
records of a large related table. A refused edit is explained in
`workspace-grid-error`, with `workspace-grid-error-dismiss` to put it away: the
grid explains its own refusals rather than writing to `workspace-error`, because
each belongs to the cell it names and is cleared only by a later attempt on that
cell or by that button. Inside the grid the identifiers are Tabulator's own: a
row carries `data-record-id` with its Record ID, and a cell carries
`tabulator-field` with its column name, so a cell is addressed by record and
column rather than by position. Editing a cell is retried as a whole
interaction, because a grid that renders rows as they are needed can move an
element under a click.

An editor opens on a **double** click, not a single one: a single click selects
the cell, because a drag from it selects a range. The fill handle on the corner
of the selection is `workspace-fill-handle`; drag it with `page.mouse` rather
than `dragTo`, so the drag passes over the cells between. A gesture refused
before anything is sent (a paste past the last record, a disjoint selection, the
Record ID as a target) reports into `workspace-grid-error` like any other
refusal, against the cell the gesture started from.

`createWorkbookUpload` accepts `{ formula }` in place of a cell value, which
writes `<f>` with no `<v>`: a formula the workbook carries no calculated value
for. A spreadsheet engine will not produce that shape, which is exactly why the
fixture has to.

Import is its own section, `workspace-import`, with `workspace-import-choose`
and the hidden `workspace-import-input` behind it. A chosen file renders
`workspace-import-form`, one `workspace-import-source` per worksheet (each with
`workspace-import-source-name`, `workspace-import-selected`,
`workspace-import-name`, `workspace-import-prefix`, and
`workspace-import-padding`), then `workspace-import-run` and
`workspace-import-cancel`. `workspace-import-problem` holds the one reason Run
is held back, and `workspace-import-error` a refusal from the worker. Because
both the page and the import section render a file input, scope the workspace
one to `workspace-import-input` rather than the bare `file-input` helper.

The preview and results panels also carry accessible names, so
`getByRole("region", { name: "Results" })` works where a role-based query reads
better.

## Fixtures

PDFs are generated in memory with pdf-lib, and workbooks and presentations are
assembled from minimal OOXML parts with jszip, all in `fixtures.ts`, and
uploaded as buffers. `createWorkbookUpload` takes one fixture per worksheet, and
a worksheet may declare a `state` (`hidden` or `veryHidden`) and an Excel
`table`; workbook-level `definedNames` are passed alongside the macro-enabled
option. Those three are what the inspection report is built to describe, so the
fixtures spell them out in the parts a package reader actually reads: a `state`
attribute on the sheet entry, a table part reached through the worksheet's own
relationships, and defined names in the workbook part.
`createPresentationUpload` takes one array of run strings per slide, so a
fixture spells out how a paragraph is split across text runs, the detail the
populate engine has to stitch back together before it can see a `{{field}}`.
`createMappingUpload` writes a column mapping document verbatim, so a test can
hand the consolidate page one the engine accepts or one it refuses. Nothing
binary is checked in and no temporary files are written.
