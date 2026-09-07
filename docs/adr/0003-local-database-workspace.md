# Local database workspace

Status: Proposed (draft for agreement). The stack and framing decisions below
were each agreed on their own before being written here. Two decisions stay
deferred to the build item that needs them: the computed-column formula
language, and what the dashboard HTML export carries (static data or an inlined
query engine).

Every ConsultChimps tool so far is a stateless operation: immutable inputs in,
artifacts out, nothing kept between runs. This feature is different. It is a
stateful workspace that holds a small relational database for a project, edited
in place across sessions by a few people taking turns from a shared folder. The
registry (ADR 0001), the operations Web Worker, and the feature-completion
checklist all model stateless operations, so the first job is to say how a
stateful workspace lives beside them without bending what an operation means.

The database is SQLite, opened in the browser tab. There is no server and no
hosted service, and data never leaves the machine. Import reuses
`@consultchimps/xlsx` and `@consultchimps/tabular` (reading, inspection, column
mapping); the `Table` model is the exchange format in both directions, so the
database can also feed the existing PowerPoint populate and split operations.
The formula-preserving Excel export is a new L3 operation on the xlsx package's
OOXML model, per `packages/xlsx/ARCHITECTURE.md`. Dependencies stay Apache-2.0
compatible: HyperFormula and ExcelJS are excluded, and so is any GPL, LGPL, or
EPL package (elkjs among them).

## Decision 1: save model

The file lives in a shared folder and one person edits at a time. The save model
has to make "open the shared file, edit, put it back" the natural path, not an
error-prone ritual.

Options considered: in-place through the File System Access API with a download
fallback; OPFS autosave plus an explicit download every session; File System
Access only with no fallback.

**Decision: in-place through the File System Access API, with an OPFS autosave
mirror and a download fallback.** On Chromium the workspace writes directly back
to the shared-folder file the user opened, holding the file handle for the
session. The OPFS copy autosaves for crash recovery and is a mirror, not a
source of truth, so nobody reconciles two locations. Safari and Firefox, which
lack the API, fall back to download-and-replace. It serves the shared-folder,
one-editor model directly and degrades rather than blocking.

## Decision 2: SQLite engine

**Decision: sql.js (MIT, SQLite compiled to WebAssembly, in memory).** The
workspace loads the whole file into memory, edits it, and serializes it back to
bytes to save, which fits the load, edit, save-back shape of decision 1 exactly
and keeps the OPFS copy a genuine serialized mirror. An OPFS or File System
Access VFS engine (official sqlite-wasm, wa-sqlite) persists incrementally, but
it turns OPFS into a second live database and still needs an explicit serialize
to write back to the shared folder, so the extra machinery buys little for a
small database. Revisit if databases outgrow a comfortable in-memory size.

## Decision 3: how the workspace fits the registry

Options considered: a workspace page outside the operation registry with
stateless operations inside it; extending the registry with a new workspace
surface kind; folding the whole feature into one operation with sub-modes.

**Decision: a workspace page outside the registry, with stateless operations
inside it**, the way `/shortcuts` and the inspect chrome already sit beside the
operations. The registry keeps describing stateless operations, each still
obeying the browser-surface rules (bytes-level, no filesystem, run in the
worker). The editor is a new "workspace" category defined here: a stateful,
browser-only page that opens a file, mutates it in place, and saves it back, and
is exempt from the no-filesystem rule precisely because it is not an operation.
This keeps the rule that matters, that a card or button never offers a
capability that does not exist, and it avoids reworking ADR 0001 to model
statefulness the rest of the toolkit does not have. The File System Access API
of decision 1 lives only on this page, never in an operation.

The operations the feature contributes to the registry:

- **`db.import`**: Excel or CSV to a SQLite database (library, CLI, browser).
- **`db.export-excel`**: database to a formula-preserving workbook (library,
  CLI, browser).
- **`db.export-dashboard`**: a dashboard definition to a self-contained HTML
  file (library, CLI, browser).

The editor, relationship diagram, and dashboard builder are the workspace page
and are browser-only.

## Decision 4: editing grid

The grid has to feel as close to Excel as possible: rectangular and disjoint
range selection, a drag fill handle, and clipboard copy and paste of ranges as
TSV that round-trips with Excel. A headless table library (TanStack Table)
supplies none of that, so it would mean building a selection and clipboard
engine by hand, which is exactly the fiddly, error-prone work to avoid.

The license floor was deliberately widened to weigh this, first to MPL, then to
GPL, then to paid commercial. The finding: the batteries-rich grids
(Handsontable, AG Grid Enterprise, Kendo React, Bryntum) are pure commercial
with no usable GPL path; the only genuine dual GPL and commercial grids (DHTMLX,
Webix) are not materially better and fight a custom design system with their own
skins. So GPL earns nothing here.

**Decision: Tabulator (MIT), used through `react-tabulator`.** Within permissive
licensing it is the most complete Excel range and clipboard package: rectangular
and disjoint (ctrl-click) range selection, keyboard range extension, a
corner-drag fill, and Excel-compatible TSV clipboard, DOM-rendered so tool-kit
CSS and light or dark style the cells, and actively maintained. It satisfies
both the Excel-UX requirement and the batteries-included goal with no license
cost and no GPL split on `apps/docs`. One item to spike before building on it:
confirm the corner-drag fill produces Excel-style series fill, not only a copy.

## Decision 5: theme package

Exports carry a client's brand colours, which is a real consulting need, but the
exports are library and CLI operations, so the theme model cannot live in
`apps/docs`.

**Decision: a new runtime-neutral package `@consultchimps/theme`.** It holds the
palette (categorical, sequential, and semantic colours), light and dark, and
validation (contrast and categorical distinctness, reusing the `dataviz` skill's
method), with zero dependencies. The dashboard HTML export consumes it now, and
the Excel export and the site can consume it later without a wrong dependency
direction. Neutral placeholder palettes only are committed; a client's colours
are supplied at runtime and never enter the repo, per the repository's
no-client-references rule.

## Decision 6: charts

The dominant constraint is the self-contained offline HTML export, which must
ship as a single file that renders without a network. The decision-maker ruled
out hand-rolled SVG (error-prone) and asked for a library, with
`@tanstack/charts` explicitly in scope despite its maturity.

Verified facts (September 2026): the mature, lower-risk choice is Recharts (MIT,
React, a documented static-SVG export path via `renderToStaticMarkup`).
`@tanstack/charts` is MIT and architecturally the best fit (framework-agnostic,
SVG server-side rendering as a headline feature, so the export generator emits
static SVG with no React runtime, and CSS-variable theming), but it is pre-1.0
(0.16, self-described as alpha) with structural breaking changes between minor
releases.

**Decision: `@tanstack/charts`, adopted now despite its pre-1.0 status, with the
risk contained.** Its architecture removes the React-in-the-generator cost that
every other library carries, and its SVG SSR produces the runtime-free offline
export this feature needs. To bound the alpha risk: pin an exact `0.16.x`, and
put all chart construction behind one thin adapter module so a breaking minor
bump touches a single file rather than every dashboard. Reassess at 1.0. KPI
stat tiles are plain JSX, not a chart type.

## Decision 7: relationship-diagram rendering

The diagram is an in-app view only, not an exported artifact, so there is no
offline or static-SVG constraint here, and per the no-hand-rolled-SVG preference
it is library-based.

**Decision: React Flow (`@xyflow/react`, MIT) for the node and edge rendering,
with `@dagrejs/dagre` (MIT) for layout.** React Flow gives a themeable,
pan-and-zoom diagram of tables and foreign-key edges; `@dagrejs/dagre` computes
a non-overlapping layout. `@dagrejs/dagre` is the maintained fork; the original
`dagre` has been unmaintained since 2019, so it is not used. Both are in-app
only and never enter an export. elkjs, the other common layout engine, is
dual-licensed EPL-2.0 or GPL-3.0, both copyleft, and therefore excluded.

## Decision 8: stable record identifiers

Relationships hang off record identity, so a record needs an identifier that is
human-readable (so it means something to a consultant and can be read in a
foreign-key cell), stable (so a relationship never breaks when other fields are
edited), and usable as a foreign-key target.

**Decision: a per-table prefixed sequential identifier, always generated, for
example `CUST-0001` or `INV-0042`.** The prefix and zero-padding are
configurable per table. The identifier is assigned once at record creation and
is immutable thereafter; foreign keys reference it, and in SQLite it is a
`UNIQUE` text column (a stable key beside the internal rowid), never an editable
display field. Single-writer editing means no counter contention, and gaps after
a delete are acceptable because the identifier is stable, not dense. Tables
always use the generated identifier; existing domain codes live in ordinary
columns and are not made the key, which keeps one identity mechanism to reason
about.

## Deferred decisions

Settled when the build item that needs them is designed, so each can be
discussed with its real constraints in front of us:

- **Computed-column formula language** (build item 7): a small language that
  compiles to both a SQLite expression and an Excel formula, and its initial
  function set.
- **Dashboard export payload** (build item 10): whether the self-contained HTML
  ships static rendered data or an inlined sql.js engine for in-page filtering.
- **Glossary verbs for `CONTEXT.md`** (build item 1): the package names are
  `@consultchimps/db` and `@consultchimps/theme`; the workspace and
  record-editing verbs are still to agree.

## Build list

Each item is independently designable and buildable, in order. The decision it
carries, if any, is noted.

1. **Package scaffolds and glossary.** Create `@consultchimps/db` (schema model
   types, the sql.js wrapper, the `Table` bridge to `@consultchimps/tabular`)
   and `@consultchimps/theme` (palette model and validation, neutral placeholder
   palettes). Add the workspace and record-editing terms to `CONTEXT.md`.
2. **Import operation.** `db.import`: Excel or CSV to SQLite, reusing xlsx,
   tabular, and column mapping, on library, CLI, and browser. Registry entry.
3. **Schema and relationships model.** Tables, columns, types, foreign keys, and
   the per-table prefixed stable identifier (Decision 8), persisted in the
   database file, in the library.
4. **Workspace shell.** The browser page that opens or creates a database
   through the File System Access API, autosaves to OPFS, and falls back to
   download, outside the registry. Adds the workspace completion checklist.
5. **Record grid.** Excel-like editing (range and disjoint selection, fill,
   clipboard), foreign-key pickers, validation, and undo, on Tabulator through
   `react-tabulator`. Spike the series-fill behaviour first.
6. **Relationship diagram.** A pan-and-zoom view of tables and foreign keys, on
   React Flow and dagre.
7. **Computed columns.** The formula language that compiles to a SQLite
   expression and an Excel formula. (Deferred decision.)
8. **Formula-preserving Excel export.** `db.export-excel`, a new xlsx L3
   operation that emits live formulas with no hardcoded results, on library,
   CLI, and browser.
9. **Dashboards.** A KPI, bar, line, and pie builder in the workspace, on
   `@tanstack/charts` behind a thin adapter, themed from `@consultchimps/theme`.
10. **Dashboard HTML export.** `db.export-dashboard`, a self-contained HTML file
    per dashboard, static SVG with no runtime, on library, CLI, and browser.
    (Deferred decision on the payload.)
11. **Bridge to existing operations.** A database table through `Table` into
    PowerPoint populate and split, so the workspace feeds the tools that already
    exist.

## Consequences

- ADR 0001 and its drift checks are unchanged. The registry gains three
  operations; the workspace is a new category this ADR defines, not a registry
  entry.
- The browser-surface rules keep their meaning: operations stay bytes-level and
  filesystem-free; only the workspace page uses the File System Access API, and
  it is not an operation.
- New dependencies, all Apache-2.0 compatible: sql.js (MIT); `tabulator-tables`
  and `react-tabulator` (MIT) for the grid; `@tanstack/charts` (MIT, pre-1.0 at
  0.16 and self-described as alpha, pinned to an exact version and isolated
  behind one adapter, reassessed at 1.0); `@xyflow/react` and `@dagrejs/dagre`
  (MIT) for the diagram, in-app only. A new first-party package
  `@consultchimps/theme` (zero dependency). The grid license floor was widened
  to MPL, GPL, and commercial and the permissive choice still won, so nothing on
  `apps/docs` relicenses; copyleft grids and elkjs (dual EPL-2.0 or GPL-3.0) are
  excluded.
- Because `@tanstack/charts` is framework-agnostic with SVG server-side
  rendering, the dashboard-export generator produces static SVG without bundling
  React into the exported file, so the offline dashboard carries no JavaScript
  runtime.
- Theming ships neutral placeholder palettes only; client colours are runtime
  input and never committed.
- The feature-completion checklist applies to the three operations as written.
  The workspace page needs its own short checklist, added when build item 4
  lands, covering save, autosave, and the download fallback rather than
  artifacts.
- Editing is single-writer by design. The workspace does not attempt concurrent
  multi-user editing; the shared-folder, one-editor-at-a-time model is a stated
  constraint, not a limitation to remove later.
