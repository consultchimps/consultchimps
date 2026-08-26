# Feature completion

This document defines what "complete" means for an operation surface in the tool
registry (`apps/docs/src/lib/tools.ts`, per
[ADR 0001](../adr/0001-feature-registry-and-drift-checks.md)).

**Binding rule:** a pull request that flips any surface's status to `works` must
include the applicable checklists below in its description, with evidence per
item — the test file, docs page, or command that demonstrates it. A flip without
the demonstrated checklist is not mergeable.

## Every surface

- [ ] Deterministic output for identical inputs and options, covered by a test.
- [ ] Success returns a structured `OperationResult`; expected failures throw
      `ConsultChimpsError` with stable, namespaced codes. Both paths are tested.
- [ ] Operations that can run long on realistic inputs support `onProgress` and
      `signal` (cancellation).
- [ ] A Changeset exists for every published-package change.
- [ ] The registry entry's status matches reality and `pnpm check` passes,
      including both drift checks.

## CLI surface (`cli: "works"`)

- [ ] Command wired with explicit option mapping, short and long flags where
      conventional, realistic help examples, and a result-shape description.
- [ ] Human-readable results on stdout, progress and diagnostics on stderr,
      `--json` envelope kept clean.
- [ ] The CLI reference page documents the command (verified by
      `scripts/check-cli-reference.ts`).
- [ ] CLI tests execute the built `dist` entry point.

## Library surface (`library: "works"`)

- [ ] Public API exported with explicit typed options and result interfaces; no
      internal types leak into the published `.d.ts`.
- [ ] The libraries guide documents the API with an executable example.

## Browser surface (`browser: { status: "works", href }`)

- [ ] Runs fully client-side on a bytes-level API — no filesystem access, files
      never leave the machine; heavy work runs in a Web Worker.
- [ ] The tool page exists at the registry `href`; cards, sub-bar tab, and the
      guide's "Try … online" button light up from the registry entry alone
      (verified by `scripts/check-registry-site.ts`).
- [ ] Playwright e2e coverage: the navigation spec lists the tool, and a
      functional spec exercises the page's happy path.
- [ ] Multi-file downloads offer the bundled zip alongside individual files.
- [ ] The operation's guide page has a section for the online tool.
