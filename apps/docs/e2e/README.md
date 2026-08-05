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
  reporting a column the workbook does not have, and refusing a file that is not
  a workbook.
- `excel-merge.spec.ts` — merging two workbooks into one, reordering and
  removing sources, and downloading the result.
- `tools-navigation.spec.ts` — the `/tools` index, the sub-bar tabs, and the
  "Try it online" button each guide gains from the tool registry.

Every downloaded PDF is checked for the `%PDF-` header and every downloaded
workbook for the `PK` ZIP header, so a tool that "finishes" while producing
empty or corrupt bytes fails the suite.

## Selectors

Address the tool pages through the `data-testid` attributes the shared tool
shell renders, not through heading text. The stable identifiers are:

| Identifier                            | Element                                 |
| ------------------------------------- | --------------------------------------- |
| `file-picker` / `file-input`          | the drop zone and its file input        |
| `source-summary`                      | the chosen single input's name and size |
| `source-list` / `source-item`         | the ordered list of merge inputs        |
| `preview-section` / `preview-error`   | the plan preview and its failure text   |
| `planned-outputs`                     | the planned output names                |
| `run-button` / `cancel-button`        | the run controls                        |
| `progress-report`                     | the progress bar and its labels         |
| `results-section`                     | the Results region                      |
| `artifact-list` / `artifact-item`     | the produced outputs                    |
| `artifact-name` / `artifact-download` | one output's name and Download button   |
| `archive-download`                    | "Download all (.zip)"                   |
| `result-message` / `failure-message`  | the outcome text, by outcome            |

The Excel split page adds `column-select`, `column-input`, and one identifier
per advanced control (`prefix-input`, `sheet-input`, `table-input`,
`range-input`, `header-row-input`, `include-blank-checkbox`,
`include-hidden-checkbox`, `preserve-workbook-checkbox`, `values-checkbox`).

The preview and results panels also carry accessible names, so
`getByRole("region", { name: "Results" })` works where a role-based query reads
better.

## Fixtures

PDFs are generated in memory with pdf-lib and workbooks are assembled from
minimal OOXML parts with jszip, both in `fixtures.ts`, and uploaded as buffers.
Nothing binary is checked in and no temporary files are written.
