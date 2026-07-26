# chimpcons

Composable, local-first operations tools for consultants.

`chimpcons` turns common document and data chores into deterministic TypeScript
modules that can run from a library, CLI, desktop application, server, or
automation. AI may assist with ambiguous work later, but the core tools do not
require Codex or any hosted service.

## Initial tools

| Tool                   | Command                        | Status  |
| ---------------------- | ------------------------------ | ------- |
| Excel consolidation    | `chimpcons sheets consolidate` | Working |
| PDF page splitting     | `chimpcons pdf split`          | Working |
| PDF document merging   | `chimpcons pdf merge`          | Working |
| Dataset inspection     | `chimpcons data inspect`       | Planned |
| Dataset reconciliation | `chimpcons data reconcile`     | Planned |

## Packages

| Package              | Responsibility                                   |
| -------------------- | ------------------------------------------------ |
| `@chimpcons/core`    | Shared errors, artifacts, and operation results  |
| `@chimpcons/files`   | Input discovery and safe output-path handling    |
| `@chimpcons/tabular` | Runtime-neutral table model and union operations |
| `@chimpcons/xlsx`    | Excel workbook input and output                  |
| `@chimpcons/pdf`     | PDF split and merge operations                   |
| `chimpcons`          | Command-line interface                           |

## Quick start

Requirements:

- Node.js 20.19 or newer
- pnpm 11

```bash
pnpm install
pnpm build

pnpm chimpcons sheets consolidate "inputs/**/*.xlsx" \
  --output outputs/consolidated.xlsx

pnpm chimpcons pdf split report.pdf --output outputs/pages

pnpm chimpcons pdf merge "inputs/**/*.pdf" \
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

See [CONTRIBUTING.md](CONTRIBUTING.md) for the contribution workflow and
[SECURITY.md](SECURITY.md) for reporting vulnerabilities.

## License

Apache-2.0
