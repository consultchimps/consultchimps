---
"@consultchimps/theme": minor
---

Add the `@consultchimps/theme` package: a runtime-neutral, zero-dependency
palette model with categorical, sequential, and semantic colours in light and
dark, and a small API to resolve a colour role to a value for a mode. It
includes a validation pass that returns structured results (never throws) for
WCAG contrast and categorical distinctness, reusing the `dataviz` method, and
ships one neutral placeholder palette. Client colours are supplied at runtime
and never committed.
