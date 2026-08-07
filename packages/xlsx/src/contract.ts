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
 * Only `split` is populated in Phase 1, and its cells encode the POST-Phase-1
 * intent: once the split runs on the L1 model, the invariant pass named in
 * `DeleteRowsReport.adjusted` fixes dependent references natively, so those
 * structures are declared `fix` even though the corpus currently pins them as
 * broken (`test/corpus/tier1-gaps.corpus.test.ts` carries the matching
 * `it.fails` pair for each one).
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
    // Tier-1 gap: the calculation chain must not reference cells a split deleted
    "calc-chain": "fix",
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
  // Phase 1b rebuilds merge as a part-level transplant; declaring cells before
  // that work would pin behavior the corpus already records as lossy.
  merge: {},
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
  "pivot-tables":
    "Tier-1 gap: pivot-cache records of another group must not reach a group's split output - strip-warn versus fix is open.",
  "external-links":
    "No corpus fixture yet; the generator has no external-link part.",
  "formulas-array":
    "pins: a single array formula anywhere on the sheet withdraws table compaction - whether Phase 1 fixes the array ref or keeps warning is open.",
};
