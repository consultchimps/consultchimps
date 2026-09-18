# @consultchimps/theme

## 0.2.0

### Minor Changes

- 4bc99d2: Raise the supported Node.js floor from 22.0.0 to 22.14.0. Continuous
  integration now tests that exact version: pnpm 11 requires 22.13, and
  better-sqlite3 13 crashes on the 22.13 patch line, so 22.14.0 is the lowest
  release on which every published package works.
- f72556d: Add the `@consultchimps/theme` package: a runtime-neutral,
  zero-dependency palette model with categorical, sequential, and semantic
  colours in light and dark, and a small API to resolve a colour role to a value
  for a mode. It includes a validation pass that returns structured results
  (never throws) for WCAG contrast and categorical distinctness, reusing the
  `dataviz` method, and ships one neutral placeholder palette. Client colours
  are supplied at runtime and never committed.
