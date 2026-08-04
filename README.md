# ConsultChimps

Composable, local-first operations tools for consultants.

`consultchimps` turns common document and data chores into deterministic
TypeScript modules that can run from a library, CLI, desktop application,
server, or automation. AI may assist with ambiguous work later, but the core
tools do not require Codex or any hosted service.

Documentation:
[consultchimps.github.io/consultchimps](https://consultchimps.github.io/consultchimps/)

## Initial tools

| Tool                   | Command                            | Status  |
| ---------------------- | ---------------------------------- | ------- |
| Excel consolidation    | `consultchimps sheets consolidate` | Working |
| Excel split by column  | `consultchimps sheets split`       | Working |
| PowerPoint population  | `consultchimps pptx populate`      | Working |
| PDF page splitting     | `consultchimps pdf split`          | Working |
| PDF document merging   | `consultchimps pdf merge`          | Working |
| Dataset inspection     | `consultchimps data inspect`       | Planned |
| Dataset reconciliation | `consultchimps data reconcile`     | Planned |

## Packages

| Package                   | Responsibility                                   |
| ------------------------- | ------------------------------------------------ |
| `@consultchimps/core`     | Shared errors, artifacts, and operation results  |
| `@consultchimps/files`    | Input discovery and safe output-path handling    |
| `@consultchimps/tabular`  | Runtime-neutral table model and union operations |
| `@consultchimps/xlsx`     | Excel workbook input and output                  |
| `@consultchimps/pptx`     | PowerPoint template inspection and population    |
| `@consultchimps/pdf`      | PDF split and merge operations                   |
| `@consultchimps/messages` | Plain-language rendering of results and errors   |
| `consultchimps`           | Command-line interface                           |

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

pnpm consultchimps sheets split clients.xlsx \
  --range ClientRange \
  --column Region \
  --output outputs/by-region

pnpm consultchimps pptx populate \
  --template profile-template.pptx \
  --data companies.xlsx \
  --sheet Companies \
  --template-slide 1 \
  --output outputs/company-profiles.pptx

pnpm consultchimps pdf split report.pdf --output outputs/pages

pnpm consultchimps pdf merge "inputs/**/*.pdf" \
  --output outputs/combined.pdf
```

Excel consolidation reads every visible, non-empty worksheet, unions columns by
case-insensitive header name, and adds `_source_file`, `_source_sheet`, and
`_source_row` columns. Original files are never modified.

Excel splitting creates one workbook per distinct value in a selected column.
Blank values are retained by default, filenames are portable, and all
destinations are checked before any output is written. A named Excel Table
(preferred) or a workbook named range can be selected to ignore titles, notes,
totals, and other cells outside its bounds. Table splits keep the complete
source workbook by default and change only the selected table's rows — prepare
the file exactly as you want to deliver it before splitting; use
`--no-preserve-workbook` for compact data-only outputs.

PowerPoint population reads `{{field_name}}` placeholders from one selected
template slide and creates one populated slide per nonempty Excel record. The
first slide and first worksheet are used by default, and both can be selected
explicitly. Placeholders may span adjacent PowerPoint text runs, while ordinary
text-shape formatting is retained. The output contains only generated slides, in
worksheet order. Source presentations and workbooks are never modified.

## Design principles

To preserve every source worksheet as a separate tab, use `consultchimps sheets merge "inputs/**/*.xlsx" --output outputs/all-sheets.xlsx`. It copies worksheet formatting/layout supported by Excel and adds a visible `Sheet Index` tab recording source names and hidden/visible status. Use `--no-index` to omit the index.

- Deterministic and local-first
- Immutable inputs by default
- Small composable operations rather than one-off scripts
- Explicit provenance and structured errors
- Detailed, plain-language messages for non-technical users
- Versioned public APIs and recipes
- Optional, replaceable OCR or AI providers
- Cross-platform behavior with no shell-specific assumptions

Normal CLI output explains what the tool did, translates result counts into
plain language, lists every created file, calls out warnings, confirms that
source files were left unchanged, and suggests practical next steps. Use
`--json` when an automation needs the compact structured result instead.

## Development

```bash
pnpm install
pnpm check
```

The verification suite builds the distributable CLI and runs command-level tests
against generated Excel, PowerPoint, and PDF fixtures in addition to
package-level tests.

The Fumadocs guide lives in `apps/docs`:

```bash
pnpm docs:dev
```

## Deploy the guide

The documentation site deploys to GitHub Pages. The `Docs` workflow
(`.github/workflows/docs.yml`) runs on every push to `main`: it builds the
Next.js site as a static export and publishes `apps/docs/out` with the official
Pages actions.

One-time repository setup: under **Settings → Pages**, set **Source** to
**GitHub Actions**.

The workflow derives the public URL and base path from the repository name, so
forks deploy without configuration. No secrets or environment variables are
required.

See [CONTRIBUTING.md](CONTRIBUTING.md) for the contribution workflow and
[SECURITY.md](SECURITY.md) for reporting vulnerabilities.

## License

Apache-2.0
