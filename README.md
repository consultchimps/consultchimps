# ConsultChimps

Composable, local-first operations tools for consultants.

`consultchimps` turns common document and data chores into deterministic
TypeScript modules that can run from a library, CLI, desktop application,
server, or automation. AI may assist with ambiguous work later, but the core
tools do not require Codex or any hosted service.

## Initial tools

| Tool                   | Command                            | Status  |
| ---------------------- | ---------------------------------- | ------- |
| Excel consolidation    | `consultchimps sheets consolidate` | Working |
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

## Quick start

Requirements:

- Node.js 24 or newer
- pnpm 11

```bash
pnpm install
pnpm build

pnpm consultchimps sheets consolidate "inputs/**/*.xlsx" \
  --output outputs/consolidated.xlsx

pnpm consultchimps pdf split report.pdf --output outputs/pages

pnpm consultchimps pdf merge "inputs/**/*.pdf" \
  --output outputs/combined.pdf
```

Excel consolidation reads every visible, non-empty worksheet, unions columns by
case-insensitive header name, and adds `_source_file`, `_source_sheet`, and
`_source_row` columns. Original files are never modified.

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

The Fumadocs guide lives in `apps/docs`:

```bash
pnpm docs:dev
```

See [CONTRIBUTING.md](CONTRIBUTING.md) for the contribution workflow and
[SECURITY.md](SECURITY.md) for reporting vulnerabilities.

## License

Apache-2.0
