# consultchimps

Local-first spreadsheet, PowerPoint, and PDF operations for consultants.

Requires Node.js 24.

## Run

```bash
npx consultchimps@latest --help
```

Or keep the command available globally:

```bash
npm install --global consultchimps
```

## Commands

```bash
consultchimps sheets consolidate "inputs/**/*.xlsx" \
  --output outputs/consolidated.xlsx

consultchimps sheets merge "inputs/**/*.xlsx" \
  --output outputs/all-sheets.xlsx

consultchimps sheets split clients.xlsx \
  --column Region \
  --output outputs/by-region

consultchimps sheets split clients.xlsx \
  --table ClientData \
  --column Region \
  --output outputs/by-region

consultchimps sheets split clients.xlsx \
  --table ClientData \
  --column Region \
  --preserve-workbook \
  --output outputs/by-region

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
option.

By default, every command gives a detailed, plain-language explanation of what
happened. It translates result counts, lists every created file, reports
warnings, confirms source-file safety, and suggests next steps. Add `--json`
before the command when an automation needs the structured result instead:

```bash
consultchimps --json pdf split report.pdf --output outputs/pages
```

Reusable TypeScript APIs are published separately under the `@consultchimps/*`
scope.
