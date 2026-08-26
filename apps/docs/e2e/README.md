# In-browser tool smoke tests

Playwright drives the statically exported site — not the dev server — so these
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

- `pdf-split.spec.ts` — splitting a two-page PDF into zero-padded page files,
  downloading one of them, and refusing a file that is not a PDF.
- `pdf-merge.spec.ts` — merging two single-page PDFs into `combined.pdf` and
  downloading the result.
- `excel-split.spec.ts` — detecting a workbook's column headers, splitting on
  one of them into a workbook per distinct value, downloading one of them,
  reading a downloaded workbook back to confirm it is a complete copy of the
  source, reporting a column the workbook does not have, and refusing a file
  that is not a workbook.
- `excel-merge.spec.ts` — merging two workbooks into one, reordering and
  removing sources, and downloading the result.
- `excel-consolidate.spec.ts` — stacking two workbooks whose headers drifted
  apart into one table, downloading the result, and checking that the "Normalize
  headers" and "Add source columns" checkboxes change the columns the finished
  workbook holds.
- `pptx.spec.ts` — populating a template slide from workbook records into one
  deck, naming the output, downloading it, reporting a placeholder no column
  feeds, refusing a template that is not a presentation, and inspecting a
  template's placeholders with their occurrence counts on the chosen slide.
- `tools-navigation.spec.ts` — the `/tools` index, the sub-bar tabs, the
  tool-named "Try ... online" button each guide gains from the tool registry,
  and the single button a guide shared by two operations offers.

Every downloaded PDF is checked for the `%PDF-` header and every downloaded
workbook or presentation for the `PK` ZIP header — both `.xlsx` and `.pptx` are
ZIP packages — so a tool that "finishes" while producing empty or corrupt bytes
fails the suite. `readWorkbookDownload` goes further and opens a downloaded
workbook with jszip, resolving each worksheet through the workbook's own
relationships, so a test can assert which worksheets and which rows reached the
user.

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

The PowerPoint populate page takes two files, so it wraps each picker in its own
section: `template-section` (with `template-summary`) and `records-section`
(with `records-summary` and the advanced controls `worksheet-input`,
`header-row-input`, `template-slide-input`, `output-name-input`). Because both
sections render a `file-input`, always scope the input to its section on that
page rather than using the bare `file-input` helper. The PowerPoint inspect page
has a single `source-section` (with `source-summary` and `template-slide-input`)
and reports into `inspection-section`, which renders `placeholder-list` with one
`placeholder-item` per placeholder — each carrying a `placeholder-name` and its
occurrence count — plus `inspection-warnings` holding one `inspection-warning`
per condition that would make a populate refuse the template, or
`inspection-error` when the template cannot be read. That page has no Run
button: choosing a template inspects it after the usual preview debounce.

Both PowerPoint pages reject a slide or row number that is not a whole number
counted from 1 rather than falling back to a default. The offending field
renders `<field>-error` — `template-slide-input-error`, `header-row-input-error`
— and the task is withdrawn: the populate page also lists the messages in
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
ms preview debounce, so they are safe to assert.

The preview and results panels also carry accessible names, so
`getByRole("region", { name: "Results" })` works where a role-based query reads
better.

## Fixtures

PDFs are generated in memory with pdf-lib, and workbooks and presentations are
assembled from minimal OOXML parts with jszip, all in `fixtures.ts`, and
uploaded as buffers. `createPresentationUpload` takes one array of run strings
per slide, so a fixture spells out how a paragraph is split across text runs —
the detail the populate engine has to stitch back together before it can see a
`{{field}}`. Nothing binary is checked in and no temporary files are written.
