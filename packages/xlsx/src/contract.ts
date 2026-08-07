/**
 * L4 seam - the contract table. See packages/xlsx/ARCHITECTURE.md.
 *
 * For every tracked workbook structure and every operation, the behavior this
 * package promises. The table is data so that tests can walk it: the
 * conformance corpus exercises the declared cells, and
 * `test/contract.test.ts` reports every structure that has no declared cell.
 *
 * A cell is deliberately ABSENT when the intended behavior is genuinely
 * undecided. Absence is the debt ledger; the completeness test asserts the
 * exact missing set, so shrinking it is always a deliberate edit rather than a
 * silent drift.
 */

/** What an operation promises to do to a structure it encounters. */
export type ContractBehavior =
  /** Carried through untouched and still semantically valid. */
  | "preserve"
  /** Rewritten so it remains valid after the edit. */
  | "fix"
  /** Removed from the output, with a warning in the result. */
  | "strip-warn"
  /** Stable `ConsultChimpsError` before any output is produced. */
  | "refuse";

/**
 * The structures tracked by the contract, from ARCHITECTURE.md's initial set.
 *
 * The architecture document lists "Excel Tables (incl. totals rows)" as one
 * structure. The corpus pins the two separately, because the totals row is the
 * one place where the bindings visibly disagree (the range binding deletes it,
 * the table binding keeps and resizes it), so they get separate cells here.
 */
export const TRACKED_STRUCTURES = [
  "merged-cells",
  "conditional-formatting",
  "data-validation",
  "hyperlinks",
  "comments",
  "drawings-charts",
  "defined-names",
  "excel-tables",
  "excel-table-totals-row",
  "pivot-tables",
  "calc-chain",
  "shared-strings",
  "styles-number-formats",
  "vba-project",
  "external-links",
  // The six formula flavours, tracked separately because the engine already
  // treats them differently: structured references survive relocation, A1
  // references do not, and array formulas withdraw table compaction entirely.
  "formulas-cached",
  "formulas-uncached",
  "formulas-shared",
  "formulas-array",
  "formulas-a1",
  "formulas-structured-ref",
] as const;

export type Structure = (typeof TRACKED_STRUCTURES)[number];

/** The operations the package exposes, per ARCHITECTURE.md's L3. */
export const OPERATIONS = [
  "split",
  "merge",
  "consolidate",
  "values",
  "describe",
] as const;

export type Operation = (typeof OPERATIONS)[number];

/**
 * The declared behaviors.
 *
 * `split` cells encode the POST-Phase-1 intent: once the split runs on the L1
 * model, the invariant pass named in `DeleteRowsReport.adjusted` fixes
 * dependent references natively, so those structures are declared `fix` even
 * though the corpus currently pins them as broken
 * (`test/corpus/tier1-gaps.corpus.test.ts` carries the matching `it.fails`
 * pair for each one).
 *
 * `merge` cells describe the Phase-1b transplant in `src/merge/`, which is
 * already the shipped behavior: `merge-consolidate.corpus.test.ts` holds every
 * one of them up today. `preserve` there means the structure is copied with
 * its worksheet part and never rewritten; `fix` means the merge had to repair
 * something a second workbook made ambiguous - a name two inputs both claim,
 * or an index into a table that is per-workbook.
 *
 * Every cell cites the corpus test that holds it up. `%s` in a cited name is
 * the shape parameter of an `it.each(SHAPES)` run.
 */
export const CONTRACT: Record<
  Operation,
  Partial<Record<Structure, ContractBehavior>>
> = {
  split: {
    // Tier-1 gap: merged ranges must follow the rows they cover
    "merged-cells": "fix",
    // Tier-1 gap: conditional-formatting sqref must shrink with the rows it covers
    "conditional-formatting": "fix",
    // Tier-1 gap: data-validation sqref must shrink with the rows it covers
    "data-validation": "fix",
    // Tier-1 gap: hyperlinks must follow the row they decorate
    hyperlinks: "fix",
    // pins: the %s binding copies every source package part into every output
    // (the part survives; its VML <x:Row> anchor still has to follow the row,
    // which is why this is `fix` and not `preserve` - see corpus/README.md)
    comments: "fix",
    // pins: structured-reference formulas let the table binding compact its table part
    "excel-tables": "fix",
    // Tier-1 fix: the calculation chain must not reference cells a split deleted
    "calc-chain": "fix",
    // Tier-1 fix: pivot caches carried other groups' rows into every output;
    // decided as strip-warn (rebuildable in Excel, never silently leaked)
    "pivot-tables": "strip-warn",
    // invariant: parts the %s split does not filter pass through byte-identical
    "shared-strings": "preserve",
    // invariant: parts the %s split does not filter pass through byte-identical
    "styles-number-formats": "preserve",
    // invariant: the %s binding round-trips a macro workbook as .xlsm
    "vba-project": "preserve",
    // invariant (128a310): the %s binding renumbers surviving rows and their cell references
    "formulas-cached": "preserve",
    // invariant: a values-only %s split bakes cached values and removes the calculation chain
    // (the uncached Summary!B4 is reported, never silently invented; a plain
    // split carries the formula through unchanged)
    "formulas-uncached": "preserve",
    // Tier-1 gap: shared-formula ranges must shrink with the rows they span
    "formulas-shared": "fix",
    // Tier-1 gap: A1 formulas on a plain worksheet must be rewritten when their row moves
    "formulas-a1": "fix",
    // pins: structured-reference formulas let the table binding compact its table part
    // (row-relative by construction, so compaction cannot invalidate them)
    "formulas-structured-ref": "preserve",

    // ABSENT, deliberately - see UNDECIDED_SPLIT_STRUCTURES below.
  },
  merge: {
    // invariant: merge keeps merged ranges, hyperlinks, comments and formulas
    // (worksheet-local geometry; a part-level copy never moves a cell)
    "merged-cells": "preserve",
    // invariant: merge carries conditional formatting and data validation on a
    // %s-shape input (the rule travels; the dxf it names is copied beside it)
    "conditional-formatting": "preserve",
    // invariant: merge carries conditional formatting and data validation on a %s-shape input
    "data-validation": "preserve",
    // invariant: a transplanted worksheet keeps its own dependents and comment
    // part (the external target survives; only its relationship id is renumbered)
    hyperlinks: "preserve",
    // invariant: a transplanted worksheet keeps its own dependents and comment part
    comments: "preserve",
    // invariant: a transplanted worksheet keeps its own dependents and comment
    // part - the legacy VML drawing that draws a comment box is copied with it
    "drawings-charts": "preserve",
    // invariant: merge carries defined names and renames a workbook-scoped
    // collision (first wins; a later duplicate takes a suffix, sheet-scoped
    // names follow their sheet's new index)
    "defined-names": "fix",
    // invariant: merge carries Excel Table parts and renames a %s-shape name
    // collision (table names and ids must be unique across one workbook)
    "excel-tables": "fix",
    // invariant: merge carries Excel Table parts and renames a %s-shape name
    // collision - the totals row is worksheet content plus `totalsRowCount`,
    // and a merge moves neither, so both bindings agree here
    "excel-table-totals-row": "preserve",
    // invariant: merge strips pivot tables and their caches, and says so
    // (same Tier-1 policy as the split: a cache is a private copy of rows)
    "pivot-tables": "strip-warn",
    // invariant: merge keeps one shared string table and remaps every
    // transplanted index - the chain is a derived index keyed by sheet id and
    // every transplanted sheet takes a new one, so the repair that keeps the
    // output valid is to drop it and ask for a full recalculation on open
    "calc-chain": "fix",
    // invariant: merge keeps one shared string table and remaps every transplanted index
    "shared-strings": "fix",
    // invariant: merge copies a %s-shape input's number format with the cells
    // that use it (referenced entries are copied, resolved and deduplicated)
    "styles-number-formats": "fix",
    // invariant: a macro project travels only into an .xlsm output - carried
    // from a single macro-bearing input into a macro-enabled output name, and
    // removed with a warning in every other case
    "vba-project": "strip-warn",
    // invariant: merge keeps merged ranges, hyperlinks, comments and formulas
    "formulas-cached": "preserve",
    // invariant: merge carries shared, array and uncached formulas verbatim
    "formulas-uncached": "preserve",
    // invariant: merge carries shared, array and uncached formulas verbatim
    // (`si` and `ref` are worksheet-local, so a copy cannot invalidate them)
    "formulas-shared": "preserve",
    // invariant: merge carries shared, array and uncached formulas verbatim
    "formulas-array": "preserve",
    // invariant: cross-sheet formulas follow a worksheet renamed by a collision
    "formulas-a1": "fix",
    // invariant: merge carries Excel Table parts and renames a %s-shape name
    // collision (a structured reference names its table, so it is rewritten
    // with it)
    "formulas-structured-ref": "fix",

    // ABSENT, deliberately - see UNDECIDED_MERGE_STRUCTURES below.
  },
  // Phase "Later" - table-aware consolidate.
  consolidate: {},
  // `values` is a policy on other operations today, not a standalone L3
  // operation; it earns its own column when it becomes one.
  values: {},
  // `describe` does not exist yet.
  describe: {},
};

/**
 * Why each split cell is still absent. Kept beside the table so the debt is
 * readable without cross-referencing the completeness test.
 */
export const UNDECIDED_SPLIT_STRUCTURES: Readonly<Record<string, string>> = {
  "drawings-charts":
    "No corpus fixture yet: the generator only emits the legacy VML drawing that belongs to a comment.",
  "defined-names":
    "Dependent on deleted rows, but DeleteRowsReport.adjusted does not cover them and no split test pins them.",
  "excel-table-totals-row":
    "The bindings disagree (pins: the range binding deletes the footer block, the table binding keeps it); one cell cannot state both.",
  "external-links":
    "No corpus fixture yet; the generator has no external-link part.",
  "formulas-array":
    "pins: a single array formula anywhere on the sheet withdraws table compaction - whether Phase 1 fixes the array ref or keeps warning is open.",
};

/**
 * Why each merge cell is still absent.
 *
 * The transplant already removes external links and warns about it - two
 * workbooks' `externalReferences` lists are addressed positionally and cannot
 * be interleaved - but a declared cell must be held up by the corpus, and the
 * generator has no external-link part to build one from. The cell lands with
 * the fixture.
 */
export const UNDECIDED_MERGE_STRUCTURES: Readonly<Record<string, string>> = {
  "external-links":
    "No corpus fixture yet; the generator has no external-link part, and a cell the corpus cannot exercise is not a contract.",
};
