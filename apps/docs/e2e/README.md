# In-browser tool smoke tests

Playwright drives the statically exported site — not the dev server — so these
tests exercise the same bundles GitHub Pages serves, including the lazily
imported PDF engine and the blob download path.

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
- `tools-navigation.spec.ts` — the `/tools` index and the sub-bar tabs.

Each downloaded artifact is checked for the `%PDF-` header, so a tool that
"finishes" while producing empty or corrupt bytes fails the suite.

## Fixtures

PDFs are generated in memory with pdf-lib in `fixtures.ts` and uploaded as
buffers. Nothing binary is checked in and no temporary files are written.
