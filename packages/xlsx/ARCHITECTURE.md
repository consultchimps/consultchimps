# @consultchimps/xlsx architecture

<!--
This document is the binding design for the xlsx package. Contributors —
human or agent, regardless of capability — implement against it. When code
and this document disagree, either the code is wrong or this document must
be changed in the same pull request, deliberately.
-->

## Why this architecture exists

The package grew two excellent but divergent lineages:

- an **Excel Table** lineage (explicit boundaries, structured references,
  table-part maintenance, workbook-preserving OOXML editing), and
- a **worksheet/range** lineage (header detection across sheets, tolerant value
  matching, whole-workbook outputs, staged transactional writes).

Both edit workbooks well. The problems were structural: the difference between
Tables and ranges leaked into every layer (nine interacting options, three error
codes policing illegal combinations), operations edited raw XML with regexes so
every new concern needed a new pass over text, the merge silently dropped
everything outside its library's object model, and each discovered edge case
landed as a local patch rather than a system invariant.

This architecture traps the Table/range difference in exactly one place, gives
every edge case a designated landing zone, and is deliberately shaped so that a
limited contributor — or a low-capability coding agent — cannot extend it
incorrectly without CI saying so.

## The layers

```
L5  Surfaces        file / bytes / CLI / browser — thin adapters, no logic
L4  Contract        (structure × operation) → behavior — a checked data table
L3  Operations      split / merge / consolidate / values / describe
L2  DataRegion      ONE interface; TableBinding and RangeBinding implement it
L1  Document model  rows, cells, refs, merges, names — edits keep invariants
L0  Package model   deterministic OOXML load/save (parts, rels, types)
```

Nothing above a layer reaches below the layer beneath it. Operations never touch
raw XML or JSZip; bindings never touch the filesystem; the package model never
interprets worksheet semantics.

### L0 — Package model (`src/package/`)

`WorkbookPackage`: parts, relationships, content types, deterministic
serialization (fixed DOS dates, `createFolders: false`, stable part ordering),
platform-pure `Uint8Array` I/O. This is the single owner of ZIP and part-path
concerns; the previous parallel implementations in `values-only.ts`,
`workbook-column-split.ts`, and `preserve-table-split.ts` converge here.

### L1 — Document model (`src/model/`)

Structured, lazily-parsed views over parts:

- `WorksheetModel` — rows and cells as objects, references as values, and the
  dependent structures as first-class collections: merged ranges, conditional
  formatting, data validation, hyperlinks, comment and drawing anchors,
  autoFilter.
- `WorkbookModel` — sheets with visibility, defined names, the tables registry,
  styles, shared strings, calcChain, pivot parts (enumerable).

**Defining property: edits maintain invariants natively.** Deleting rows through
the model performs, in one pass: row renumbering, cell-reference renumbering,
merged-range and sqref adjustment, table-ref resizing, and calcChain
invalidation. A bug class like "rows were renumbered but conditional-formatting
ranges were not" must be unrepresentable, because there is no way to delete rows
that bypasses the invariant pass.

Serialization writes back only parts the model actually changed; untouched parts
pass through byte-identical (the preservation guarantee the split engine
pioneered, now universal).

### L2 — DataRegion (`src/region/`)

The only place where "Excel Table" and "worksheet range" differ.

```ts
interface DataRegion {
  readonly sheet: string;
  readonly header: RowRef;
  readonly columns: readonly ColumnInfo[];
  readRecords(policy: ReadPolicy): WorksheetRecords;
  filterRows(keep: RowPredicate): RegionEditReport;
  appendRows(rows: WorksheetRecords, policy: AppendPolicy): RegionEditReport;
}
```

Exactly two implementations:

- `TableBinding` — owns totals rows, structured-reference rules, table-part
  `ref`/autoFilter maintenance.
- `RangeBinding` — owns detected/named/A1 boundaries and A1-formula safety rules
  (including the delete-without-renumbering guard inherited from the split
  engine).

One resolver owns all discovery heuristics:

```ts
type RegionSelector =
  | { table: string }
  | { range: string }              // named range or "Sheet!A1:F200"
  | { sheet: string; headerRow?: number }
  | { find: string }               // header-text search on one best sheet
  | "all-worksheets";              // header-text search across every sheet

resolveRegions(workbook: WorkbookModel, selector: RegionSelector,
               column?: string): DataRegion[];
```

Header detection (NFKC/trim/case-insensitive match, `headerRow` overrides,
table-range association) lives here and nowhere else.

### L3 — Operations (`src/operations/`)

Each operation is a short composition over regions and the model. Options are
expressed as shared, typed **policy objects** so they cannot drift apart between
operations:

- `MatchingPolicy` — strict vs. normalized value comparison.
- `BlankPolicy` — include/skip blank key values.
- `ValuesPolicy` — formulas kept, or values-only conversion.
- `VisibilityPolicy` — hidden / very-hidden sheet handling.
- `NamingPolicy` — output naming, prefixes, sanitization (core's
  `safeNameFragment`), collision suffixing.

Operations receive and return the established core contracts (`OperationPlan`,
`OperationResult`, `ByteOperationOutcome`, abort via `throwIfAborted`, progress
via `onProgress`) — unchanged.

### L4 — The contract table (`src/contract.ts` + conformance corpus)

For every workbook structure and every operation, the declared behavior:

| Behavior     | Meaning                                                   |
| ------------ | --------------------------------------------------------- |
| `preserve`   | carried through untouched and still semantically valid    |
| `fix`        | rewritten so it remains valid after the edit              |
| `strip-warn` | removed from the output, with a warning in the result     |
| `refuse`     | stable `ConsultChimpsError` before any output is produced |

The table is data, checked by tests: the conformance corpus exercises every
declared cell, and a completeness test fails when a structure × operation
combination exists in fixtures but has no declared behavior. Adding a feature
therefore _forces_ a decision for every structure — the build demands it, not a
reviewer.

Structures tracked (initial set): merged cells, conditional formatting, data
validation, hyperlinks, comments, drawings/charts, defined names, Excel Tables
(incl. totals rows), pivot tables + caches, calcChain, shared strings,
styles/number formats, VBA (`.xlsm`), external links, formulas (cached,
uncached, shared, array, A1, structured-ref).

#### The merge's removals

Preservation is the merge's default, so its four removals are the part of the
contract worth stating in prose. Each is reported as a warning in the
`OperationResult` except where noted.

| Removed        | Why a merge cannot carry it                                                                                                                                                                                                                       |
| -------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| pivot tables   | A cache is a private copy of its source rows and the `pivotCaches` registry is workbook-scoped. Same Tier-1 policy as the split.                                                                                                                  |
| external links | An external reference is addressed by its position in one workbook's `externalReferences` list, and formulas cite that position; two lists cannot interleave.                                                                                     |
| calcChain      | A derived index keyed by sheet id, and every transplanted sheet takes a new one. Dropped, and `calcPr fullCalcOnLoad` asks Excel to rebuild it. No warning: nothing authored is lost, which is why the cell reads `fix` rather than `strip-warn`. |
| VBA project    | Two `vbaProject.bin` files cannot be combined. One travels, but only when a single input carries it and the output is named `.xlsm`, because a package whose content type contradicts its file name is one Excel opens with a corruption warning. |

Two workbook-unique namespaces are repaired rather than dropped: Excel Table
names (and ids) and workbook-scoped defined names. First claim wins; a later
duplicate takes a numeric suffix, the rename is warned about, and every formula
that named it - structured references included - is rewritten to follow it.
Sheet-scoped defined names never collide and simply follow their sheet's new
index.

### L5 — Surfaces (`src/index.ts`, `src/bytes.ts`, CLI, browser)

Thin adapters only. The filesystem surface adds the staged
temp-directory/rename/rollback commit strategy (inherited from the split engine)
around byte outcomes. The public API of this package does not change during
migration; existing exports keep their signatures and error codes.

## The four guardrails

1. **Symmetry harness.** Corpus fixtures are authored in _pairs_ — the same data
   as an Excel Table and as a plain range. Operation tests run against both
   bindings automatically. A capability implemented for only one shape is a
   failing test, not a review comment.
2. **Boundary enforcement.** An import-graph test (same technique as the dist
   `node:`-import walkers) asserts that operation modules never import JSZip or
   regex-edit XML, and that model modules never import `node:fs`. Fixes
   physically cannot land in the wrong layer.
3. **Contract completeness.** Undeclared structure × operation behavior is a
   build failure (see L4).
4. **The cookbook** (below). Recipes over architecture: limited models follow
   recipes well.

## Cookbook — how to contribute

**Fix a discovered edge case.** 1) Add a fixture workbook (or extend the
generator) reproducing it, in both Table and range form when applicable. 2) Add
the failing assertion to the corpus. 3) Fix inside the _model_ (L1) if it is an
invariant, or the _binding_ (L2) if it is Table/range-specific. Never fix inside
an operation. 4) If behavior for a structure changed, update the contract cell
in the same PR.

**Add an operation.** Compose it in `src/operations/` from `resolveRegions`,
region methods, and policies. Declare a contract cell for every tracked
structure (the completeness test lists what is missing). Add corpus runs for
both bindings. Expose through file + bytes surfaces; CLI last.

**Add structure handling.** Extend the model with the structure as a first-class
collection, wire it into the invariant pass, add paired fixtures, declare its
cell for every operation.

**Move a module into a layer.** `test/boundaries.test.ts` enforces guardrail 2
over `src/**/*.ts` — source text, not `dist/`, so a misplaced import fails
before it is bundled. It carried a temporary allowlist of the legacy top-level
modules that still imported JSZip; Phase 1 emptied it, so `src/package/` is now
the only owner of ZIP concerns. The allowlist is asserted to match the offending
files _exactly_, so a new module cannot join it unnoticed. The "operations never
regex-edit XML" half of guardrail 2 landed with `src/operations/`, which the
inspection operation created: no module under it may import the XML _mutation_
helpers (`editElements`, `setAttribute`, `addAttribute`), because rewriting a
part is L0/L1 work. Reading helpers stay available — an operation may decode
text it compares against, which is why `decodeXmlText` is not on that list.

**Declare a contract cell.** `src/contract.ts` holds the L4 table as data, with
each cell citing the corpus test that holds it up. `split` cells state the
_post-Phase-1_ intent, so a cell can read `fix` while
`tier1-gaps.corpus.test.ts` still pins the gap; `merge` cells describe shipped
behavior, held up by `merge-consolidate.corpus.test.ts` today. A structure whose
behavior is genuinely undecided has **no cell**: `test/contract.test.ts` asserts
the exact missing set (with a reason per entry in `UNDECIDED_SPLIT_STRUCTURES`
and `UNDECIDED_MERGE_STRUCTURES`), which makes the debt enumerable and shrinking
it deliberate, rather than failing the build during migration. A cell the corpus
cannot exercise is not a contract: merge's `external-links` stays absent until
the generator grows an external-link fixture, even though the engine already
strips and warns about it.

**Never**: edit raw worksheet XML from an operation; add an option to one
operation that duplicates a policy; change an existing error-code value;
introduce a second implementation of anything L0–L2 owns.

## Migration plan (strangler, not rewrite)

- **Phase 0** — conformance corpus pinning _current_ behavior of both lineages
  (every existing fix becomes a named invariant test), plus expected-failure
  tests documenting the known gaps: pivot-cache contents crossing split outputs,
  stale cached aggregates in values-mode splits, calcChain left stale after row
  deletion, and dependent-range references after row compaction. Zero behavior
  change.
- **Phase 1** — L0/L1/L2 built; the all-worksheet split re-expressed on them
  behind the existing public API; symmetry harness and boundary tests active;
  Phase 0's dependent-reference expected failures flipped to passing. The
  compact rebuild modes (`preserveWorkbook: false`, named-range and worksheet
  selections) and the preserved table-selection split stay on their previous
  implementations: the first two rebuild a workbook from parsed values rather
  than editing one, and the third's contract is a refusal
  (`XLSX_SPLIT_PRESERVE_FORMULA`) rather than a repair. Both decisions are
  recorded in `src/preserve-table-split.ts` and pinned by the corpus.
- **Phase 1b** — merge rebuilt as a part-level transplant on the model (styles
  remapped, shared strings deduped, defined-name and table-name collisions
  handled). Pinned merge expectations updated deliberately. **Landed**, in
  `src/merge/`: the first input seeds the output package and later inputs'
  worksheet parts are copied into it with their dependents, so the only things
  rewritten are part paths, relationship ids, the two per-workbook index spaces
  (shared strings, styles/dxfs) and the names a collision forced to change. Its
  contract column is populated; see "The merge's removals" below.
- **Phase 2** — the remaining contract cells decided: defined names, drawings
  and charts, external links, and the totals-row asymmetry between bindings.
- **Later** — group mapping (many values → one output), multi-column split keys,
  table-aware consolidate, region-aware browser workbench chaining.

Changesets accompany any phase that alters published behavior; breaking changes
ship as minors per the 0.x convention.
