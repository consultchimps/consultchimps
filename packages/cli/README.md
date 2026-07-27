# consultchimps

Local-first spreadsheet and PDF operations for consultants.

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

consultchimps pdf split report.pdf --output outputs/pages

consultchimps pdf merge "inputs/**/*.pdf" \
  --output outputs/combined.pdf
```

Inputs are never modified. Existing outputs require the explicit `--force`
option.

Reusable TypeScript APIs are published separately under the `@consultchimps/*`
scope.
