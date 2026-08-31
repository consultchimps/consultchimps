# consultchimps

Local-first spreadsheet, PowerPoint, and PDF operations for consultants.

Requires Node.js 22 or later.

## Run

```bash
npx consultchimps@latest --help
```

Or keep the command available globally:

```bash
npm install --global consultchimps
```

## Commands

Every command, option, default, and example is documented in the
[CLI reference](https://consultchimps.github.io/consultchimps/docs/reference/cli),
which the repository checks against the executable's own help output in CI. This
README is an overview, not a catalogue: when the two disagree, the reference is
right.

| Command                 | What it does                                                 |
| ----------------------- | ------------------------------------------------------------ |
| `sheets consolidate`    | Stack rows from many workbooks into one auditable table      |
| `sheets merge`          | Copy every source worksheet into one workbook as its own tab |
| `sheets split`          | Write one workbook per distinct value in a chosen column     |
| `pptx inspect-template` | List the placeholders one template slide expects             |
| `pptx populate`         | Fill a template slide once per record from a workbook        |
| `pdf split`             | Write one zero-padded file per page                          |
| `pdf merge`             | Assemble source PDFs in resolved order into one document     |

Run `consultchimps <group> <command> --help` for the options of any one of them.
A few representative invocations:

```bash
consultchimps sheets consolidate "inputs/**/*.xlsx" \
  --output outputs/consolidated.xlsx

consultchimps sheets merge "inputs/**/*.xlsx" \
  --values \
  --output outputs/all-sheets.xlsx

consultchimps sheets split clients.xlsx \
  --column Region \
  --output outputs/by-region

consultchimps pptx inspect-template profile-template.pptx \
  --template-slide 1

consultchimps pptx populate \
  --template profile-template.pptx \
  --data companies.xlsx \
  --sheet Companies \
  --template-slide 1 \
  --output outputs/company-profiles.pptx

consultchimps pdf split report.pdf --output outputs/pages

consultchimps pdf merge "inputs/**/*.pdf" \
  --output outputs/combined.pdf
```

Inputs are never modified. Existing outputs require the explicit `--force`
option. Every spreadsheet command accepts `--values`; formulas are replaced by
their stored results without removing cell formatting. Consolidation and compact
data-only splits already write values, so the flag makes that intent explicit.

By default, every command gives a detailed, plain-language explanation of what
happened. It translates result counts, lists every created file, reports
warnings, confirms source-file safety, and suggests next steps. Add `--json`
before the command when an automation needs the structured result instead:

```bash
consultchimps --json pdf split report.pdf --output outputs/pages
```

`--json` prints exactly one JSON object on a single line of stdout. Success
prints `{"ok":true,"result":...}` and failure prints
`{"ok":false,"error":{"message":...,"code":...}}` while keeping the nonzero exit
code, so an automation can branch on `ok` instead of parsing message text.

Reusable TypeScript APIs are published separately under the `@consultchimps/*`
scope.
