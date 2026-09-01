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

The
[CLI reference](https://consultchimps.github.io/consultchimps/docs/reference/cli)
lists the commands and their options; CI checks that it names every command and
long option the executable exposes, so nothing goes missing from it. It does not
reproduce every example or default, and neither does this README, which is an
overview rather than a catalogue. `consultchimps <group> <command> --help` is
the executable speaking for itself and is authoritative on both.

| Command                 | What it does                                                 |
| ----------------------- | ------------------------------------------------------------ |
| `sheets inspect`        | Describe a workbook's sheets, headers, and samples           |
| `sheets consolidate`    | Stack rows from many workbooks into one auditable table      |
| `sheets merge`          | Copy every source worksheet into one workbook as its own tab |
| `sheets split`          | Write one workbook per distinct value in a chosen column     |
| `pptx inspect-template` | List the placeholders one template slide expects             |
| `pptx populate`         | Fill a template slide once per record from a workbook        |
| `pdf split`             | Write one zero-padded file per page                          |
| `pdf merge`             | Assemble source PDFs in resolved order into one document     |

A few representative invocations:

```bash
consultchimps sheets inspect clients.xlsx --hidden --samples 2

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
option. `sheets inspect` creates no output at all: it describes the workbook and
stops. Consolidation, worksheet merging, and splitting accept `--values`;
formulas are replaced by their stored results without removing cell formatting.
Consolidation and compact data-only splits already write values, so the flag
makes that intent explicit.

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
