# ConsultChimps

Composable, local-first operations tools for consultants.

`consultchimps` turns common document and data chores into deterministic
TypeScript modules that can run from a library, CLI, desktop application,
server, or automation. AI may assist with ambiguous work later, but the core
tools do not require Codex or any hosted service.

Documentation:
[consultchimps-docs.vercel.app](https://consultchimps-docs.vercel.app)

## Initial tools

| Tool                   | Command                            | Status  |
| ---------------------- | ---------------------------------- | ------- |
| Excel consolidation    | `consultchimps sheets consolidate` | Working |
| Excel split by column  | `consultchimps sheets split`       | Working |
| PDF page splitting     | `consultchimps pdf split`          | Working |
| PDF document merging   | `consultchimps pdf merge`          | Working |
| Dataset inspection     | `consultchimps data inspect`       | Planned |
| Dataset reconciliation | `consultchimps data reconcile`     | Planned |

## Packages

| Package                  | Responsibility                                   |
| ------------------------ | ------------------------------------------------ |
| `@consultchimps/core`    | Shared errors, artifacts, and operation results  |
| `@consultchimps/files`   | Input discovery and safe output-path handling    |
| `@consultchimps/tabular` | Runtime-neutral table model and union operations |
| `@consultchimps/xlsx`    | Excel workbook input and output                  |
| `@consultchimps/pdf`     | PDF split and merge operations                   |
| `consultchimps`          | Command-line interface                           |

## Install

Requirements:

- Node.js 24

Run the CLI without keeping a global installation:

```bash
npx consultchimps@latest --help
```

Or install it globally:

```bash
npm install --global consultchimps
consultchimps --help
```

Applications should install only the library boundary they need. For example,
the workbook adapter brings its internal ConsultChimps dependencies with it:

```bash
npm install @consultchimps/xlsx
```

All published packages are public on npm under the `consultchimps` name and
`@consultchimps` organization scope.

## Run from source

Development requires pnpm 11:

```bash
pnpm install
pnpm build

pnpm consultchimps sheets consolidate "inputs/**/*.xlsx" \
  --output outputs/consolidated.xlsx

pnpm consultchimps sheets split clients.xlsx \
  --column Region \
  --output outputs/by-region

pnpm consultchimps sheets split clients.xlsx \
  --table ClientData \
  --column Region \
  --output outputs/by-region

pnpm consultchimps pdf split report.pdf --output outputs/pages

pnpm consultchimps pdf merge "inputs/**/*.pdf" \
  --output outputs/combined.pdf
```

Excel consolidation reads every visible, non-empty worksheet, unions columns by
case-insensitive header name, and adds `_source_file`, `_source_sheet`, and
`_source_row` columns. Original files are never modified.

Excel splitting creates one workbook per distinct value in a selected column.
Blank values are retained by default, filenames are portable, and all
destinations are checked before any output is written. A named Excel Table can
be selected to ignore titles, notes, totals, and other cells outside its range.

## Design principles

- Deterministic and local-first
- Immutable inputs by default
- Small composable operations rather than one-off scripts
- Explicit provenance and structured errors
- Versioned public APIs and recipes
- Optional, replaceable OCR or AI providers
- Cross-platform behavior with no shell-specific assumptions

## Development

```bash
pnpm install
pnpm check
```

The verification suite builds the distributable CLI and runs command-level tests
against generated Excel and PDF fixtures in addition to package-level tests.

The Fumadocs guide lives in `apps/docs`:

```bash
pnpm docs:dev
```

## Deploy the guide

Import this repository into Vercel and set the project Root Directory to
`apps/docs`. Keep **Include source files outside of the Root Directory** enabled
so Vercel can use the workspace lockfile. Vercel then reads
`apps/docs/vercel.ts`, installs with pnpm, builds the Next.js application, and
uses Node.js 24 from `apps/docs/package.json`.

No environment variables are required for the documentation site.

See [CONTRIBUTING.md](CONTRIBUTING.md) for the contribution workflow and
[SECURITY.md](SECURITY.md) for reporting vulnerabilities.

## License

Apache-2.0
