# ADR 0001: Per-operation feature registry and drift checks

**Status:** accepted (2026-08-25)

## Context

The documentation site and the shipped toolkit drifted apart in three distinct
ways: a guide page's "Try it online" button opened a different tool than the
page primarily documented, the hand-written CLI reference missed options the CLI
actually has, and public library APIs shipped with no documentation at all. The
tool registry (`apps/docs/src/lib/tools.ts`) prevented drift between the
surfaces it modelled (cards, tabs, buttons), but its schema could not express
"this page covers two operations, only one of which runs in the browser", and
nothing verified the registry, the CLI, and the docs against each other.

## Decision

1. **One registry file in the docs app** remains the single source of truth for
   feature metadata. Feature metadata does not move into the packages, and
   packages do not export UI.
2. **One registry entry per operation** the toolkit ships (consolidate, workbook
   merge, spreadsheet split, PDF split, PDF merge, PowerPoint populate,
   PowerPoint template inspection), not one per marketing card.
3. **Each entry declares per-surface status**: for each of CLI, library, and
   browser, one of `works`, `planned`, or `none`. Site surfaces (cards, tabs,
   guide-page buttons, availability sentences) derive from this status; a
   browser button can only render for an operation whose browser surface is
   `works`.
4. **`planned` renders nothing user-visible.** The value exists so maintainers
   can distinguish "deliberately not built yet" from "forgotten"; visitors see a
   feature only when it works.
5. **Two drift checks run inside `pnpm check`** (same CI job, all platforms):
   - CLI ↔ reference: every command and option in the built CLI's `--help`
     output appears in `docs/content/docs/reference/cli.mdx`.
   - Registry ↔ site: browser surfaces only render for `works`, and every
     entry's documentation link resolves.
6. **A library-exports ↔ docs coverage check is rejected**, not deferred. No PR
   should add one unless this decision is explicitly reopened.

## Consequences

- Adding a CLI flag without updating the reference page fails CI. That friction
  is intended.
- New operations must be registered before the site can present them, which adds
  one small step to shipping a feature and removes the class of bug where a page
  implies a capability that does not exist.
- Shared library APIs (readers, table primitives, byte-level entry points) are
  treated as capabilities rather than features: they get documentation but no
  registry entry, cards, or tabs.
