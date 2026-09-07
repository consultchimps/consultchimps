# Local database workspace

Status: Proposed (draft for agreement). Decisions 1 to 3 below are settled;
decisions 4 to 7 are deferred to the build item that needs them.

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
compatible: HyperFormula and ExcelJS are excluded.

## Decision 1: save model

The file lives in a shared folder and one person edits at a time. The save model
has to make "open the shared file, edit, put it back" the natural path, not an
error-prone ritual.

Options considered:

- **In-place through the File System Access API, with a download fallback.** On
  Chromium the workspace writes directly back to the shared-folder file the user
  opened. Safari and Firefox, which lack the API, fall back to
  download-and-replace. An OPFS mirror autosaves for crash recovery regardless.
- **OPFS autosave plus explicit download.** Always autosave to origin-private
  storage; the user downloads to return the file to the shared folder each
  session. Identical in every browser, but the shared file is never edited in
  place, so every session is open, download, replace, which is easy to get wrong
  when people take turns.
- **File System Access only, no fallback.** One save path, true in-place save,
  but Safari and Firefox users cannot use the editor at all.

**Decision: in-place through the File System Access API, with an OPFS autosave
mirror and a download fallback.** It serves the shared-folder, one-editor model
directly and degrades rather than blocking. The workspace holds a file handle
for the session and writes back on save; the OPFS copy is a crash-recovery
mirror, not the source of truth, so nobody has to reconcile two locations.

Engine sub-decision, following from the above: **use sql.js (MIT, SQLite
compiled to WebAssembly, in memory)**. The workspace loads the whole file into
memory, edits it, and serializes it back to bytes to save, which fits the load,
edit, save-back shape exactly and keeps the save path simple. An OPFS or File
System Access VFS engine (official sqlite-wasm, wa-sqlite) persists
incrementally, but writing back to the shared folder still needs an explicit
serialize step, so the extra machinery buys little for a small database. Revisit
if databases outgrow comfortable in-memory size.

A consequence for decision 2: the File System Access API is filesystem access,
which the browser-surface rules forbid for operations. That is allowed here
because the workspace page is not an operation (see decision 2); the stateless
operations stay bytes-level and never touch the filesystem.

## Decision 2: how the workspace fits the registry

Options considered:

- **A workspace page outside the operation registry, plus stateless operations
  inside it.** The interactive editor is a top-level page that is not a registry
  operation, the way `/shortcuts` and the inspect chrome already are. Import,
  Excel export, and dashboard export are real registry operations with library,
  CLI, and browser surfaces. ADR 0003 names "workspace" as a category distinct
  from "operation"; ADR 0001 and its drift checks are untouched.
- **Extend the registry with a workspace surface.** Make the workspace a
  first-class registry entry with a new surface or status kind. More uniform,
  but it changes ADR 0001, the `ToolSurfaces` type, `check-registry-site`, and
  needs a workspace variant of the completion checklist.
- **One operation umbrella.** Model the whole feature as a single operation with
  sub-modes. Fewest entries, but it stretches the operation definition and the
  completion checklist past what they mean today.

**Decision: a workspace page outside the registry, with stateless operations
inside it.** The registry keeps describing stateless operations, each still
obeying the browser-surface rules (bytes-level, no filesystem, run in the
worker). The editor is a new "workspace" category defined here: a stateful,
browser-only page that opens a file, mutates it in place, and saves it back, and
is therefore exempt from the no-filesystem rule precisely because it is not an
operation. This keeps the one rule that matters, that a card or button never
offers a capability that does not exist, and avoids reworking ADR 0001 to model
statefulness the rest of the toolkit does not have.

The operations the database feature contributes to the registry:

- **`db.import`**: Excel or CSV to a SQLite database (library, CLI, browser).
- **`db.export-excel`**: database to a formula-preserving workbook (library,
  CLI, browser).
- **`db.export-dashboard`**: a dashboard definition to a self-contained HTML
  file (library, CLI, browser).

The editor, relationship diagram, and dashboard builder are the workspace page
and are browser-only.

## Decision 3: how the grid is built

Options considered:

- **A headless table library, rendered by us.** TanStack Table and TanStack
  Virtual (MIT, Apache-2.0 compatible) supply the table, editing, and
  virtualization logic; every cell is rendered with tool-kit components so the
  look stays ours. Avoids rebuilding the hard parts while keeping the site's
  aesthetic and no third-party chrome.
- **Hand-rolled.** Build the grid from tool-kit primitives with no new
  dependency: maximum control, Apache concerns moot, but foreign-key pickers,
  keyboard navigation, virtualization, and undo are a large amount of careful
  work.
- **A full grid component.** A batteries-included grid such as Glide Data Grid
  (MIT): fastest to a rich editor, but risks looking like an embedded
  third-party tool and adds a heavier bundle to a static page.

**Decision: a headless library, rendered by us.** TanStack Table plus Virtual
give the grid logic; tool-kit components give the look. This is the balance the
brief asks for, an editor that does not look like a third-party tool, without
rebuilding virtualization and edit-state management by hand.

## Deferred decisions

These are settled when the build item that needs them is designed, so each can
be discussed with its real constraints in front of us:

- **4. Computed-column DSL** (build item 7): a small formula language that
  compiles to both a SQLite expression and an Excel formula, and its initial
  function set.
- **5. Charts and dashboard export** (build items 9 and 10): hand-rolled SVG or
  a dependency, and whether the exported HTML ships static data or an inlined
  SQLite engine for in-page filtering.
- **6. Package name and glossary** (build item 1): proposed `@consultchimps/db`;
  the workspace and record-editing verbs for CONTEXT.md to be agreed.
- **7. Surface split** (per item): the editor is browser-only; import, Excel
  export, and dashboard export are library, CLI, and browser operations.

## Build list

Each item is independently designable and buildable, in order. The decision it
carries, if any, is noted.

1. **Package scaffold and glossary.** Create `@consultchimps/db` with the schema
   model types, the sql.js engine wrapper, and the `Table` bridge to
   `@consultchimps/tabular`. Add the workspace and record-editing terms to
   `CONTEXT.md`. (Decision 6.)
2. **Import operation.** `db.import`: Excel or CSV to SQLite, reusing xlsx,
   tabular, and column mapping, on library, CLI, and browser. Registry entry.
3. **Schema and relationships model.** Tables, columns, types, and foreign keys,
   persisted in the database file, in the library.
4. **Workspace shell.** The browser page that opens or creates a database
   through the File System Access API, autosaves to OPFS, and falls back to
   download, outside the registry. (Decision 1 mechanics.)
5. **Record grid.** Editing with foreign-key pickers, validation, and undo, on
   the headless library. (Decision 3.)
6. **Relationship diagram.** A read-only view of the tables and their foreign
   keys.
7. **Computed columns.** The formula DSL that compiles to a SQLite expression
   and an Excel formula. (Decision 4.)
8. **Formula-preserving Excel export.** `db.export-excel`, a new xlsx L3
   operation that emits live formulas with no hardcoded results, on library,
   CLI, and browser.
9. **Dashboards.** A KPI, bar, line, and pie builder in the workspace. (Decision
   5, charts.)
10. **Dashboard HTML export.** `db.export-dashboard`, a self-contained HTML file
    per dashboard, on library, CLI, and browser. (Decision 5, export.)
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
- A new runtime dependency enters the browser bundle: sql.js (SQLite in
  WebAssembly), plus TanStack Table and Virtual for the grid. All are Apache-2.0
  compatible.
- The feature-completion checklist applies to the three operations as written.
  The workspace page needs its own short checklist, added when build item 4
  lands, covering save, autosave, and the download fallback rather than
  artifacts.
- Editing is single-writer by design. The workspace does not attempt concurrent
  multi-user editing; the shared-folder, one-editor-at-a-time model is a stated
  constraint, not a limitation to remove later.
