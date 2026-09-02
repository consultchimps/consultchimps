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
  "unprotect",
] as const;

export type Operation = (typeof OPERATIONS)[number];

/**
 * The declared behaviors.
 *
 * `split` cells are shipped behavior: the split runs on the L1 model, whose
 * invariant pass fixes every dependent reference named in
 * `DeleteRowsReport.adjusted`, and `test/corpus/tier1-gaps.corpus.test.ts`
 * holds each one up with a passing test rather than an expected failure.
 *
 * `merge` cells describe the Phase-1b transplant in `src/merge/`, which is
 * also shipped behavior: `merge-consolidate.corpus.test.ts` holds every
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
    // Tier-1 fix: merged ranges follow the rows they cover
    "merged-cells": "fix",
    // Tier-1 fix: conditional-formatting sqref shrinks with the rows it covers
    "conditional-formatting": "fix",
    // Tier-1 fix: data-validation sqref shrinks with the rows it covers
    "data-validation": "fix",
    // Tier-1 fix: hyperlinks follow the row they decorate
    hyperlinks: "fix",
    // Tier-1 fix: a cell comment and its VML shape follow the row they annotate
    // (the part survives, and its ref and VML <x:Row> anchor move with the
    // record, which is why this is `fix` and not `preserve`)
    comments: "fix",
    // pins: structured-reference formulas let the table binding compact its table part
    "excel-tables": "fix",
    // Tier-1 fix: the calculation chain does not reference cells a split deleted
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
    // Tier-1 fix: shared-formula ranges shrink with the rows they span
    "formulas-shared": "fix",
    // Tier-1 fix: A1 formulas on a plain worksheet are rewritten when their row moves
    "formulas-a1": "fix",
    // pins: a single array formula anywhere on the sheet withdraws table compaction
    // (DECIDED IN PHASE 1: the array formula's own `ref` span and its A1 text
    // follow the rows, so the formula stays valid; what the array formula still
    // withdraws is the table-part resize, which is warned about, not silent)
    "formulas-array": "fix",
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
  /**
   * The inspection reads and writes nothing, so every tracked structure is
   * `preserve` - and uniquely, that is provable rather than argued: the column
   * is held up by asserting the input package is byte-identical after a
   * describe, so no cell here can be true while another is false.
   *
   * This is the one column where a uniform answer is the honest one. A read-only
   * operation cannot `fix` (it rewrites nothing), cannot `strip-warn` (it emits
   * no output to strip from), and does not `refuse` on any structure: the
   * refusals it does have - a workbook with no worksheets, an unknown worksheet
   * name, an out-of-range sample count, an invalid header row - are about
   * options and workbook shape, not about a tracked structure.
   *
   * Every cell is held up by `describe.corpus.test.ts`.
   */
  describe: {
    // invariant: describing a %s workbook leaves every input byte untouched
    "merged-cells": "preserve",
    // invariant: describing a %s workbook leaves every input byte untouched
    "conditional-formatting": "preserve",
    // invariant: describing a %s workbook leaves every input byte untouched
    "data-validation": "preserve",
    // invariant: describing a %s workbook leaves every input byte untouched
    hyperlinks: "preserve",
    // invariant: describing a %s workbook leaves every input byte untouched
    comments: "preserve",
    // invariant: describing a %s workbook leaves every input byte untouched
    // (the comment's legacy VML drawing is carried by the same guarantee)
    "drawings-charts": "preserve",
    // pins: the description reports defined names without altering them
    // (the workbook-scoped CorpusRange is reported; the sheet-scoped LocalNote
    // is reported too, and neither part is rewritten)
    "defined-names": "preserve",
    // pins: the description reports an Excel Table's declared headers
    // (read from the table part, so a table with no data rows still appears)
    "excel-tables": "preserve",
    // pins: a totals row is described as worksheet content and never removed
    // (the two bindings disagree about deleting it, but describe deletes
    // nothing, so this is the one operation where the cell is unambiguous)
    "excel-table-totals-row": "preserve",
    // invariant: describing a %s workbook leaves every input byte untouched
    // (a pivot cache is a private copy of rows; an inspection neither reads it
    // into the description nor disturbs it)
    "pivot-tables": "preserve",
    // invariant: describing a %s workbook leaves every input byte untouched
    "calc-chain": "preserve",
    // invariant: describing a %s workbook leaves every input byte untouched
    "shared-strings": "preserve",
    // invariant: describing a %s workbook leaves every input byte untouched
    "styles-number-formats": "preserve",
    // invariant: describing a macro workbook leaves every input byte untouched
    "vba-project": "preserve",
    // invariant: describing a %s workbook leaves every input byte untouched
    "formulas-cached": "preserve",
    // pins: a formula cell is sampled by its cached value, not recalculated
    // (an uncached formula therefore contributes no sample rather than a
    // guess - the inspection never invents a value the workbook does not hold)
    "formulas-uncached": "preserve",
    // invariant: describing a %s workbook leaves every input byte untouched
    "formulas-shared": "preserve",
    // invariant: describing a %s workbook leaves every input byte untouched
    "formulas-array": "preserve",
    // invariant: describing a %s workbook leaves every input byte untouched
    "formulas-a1": "preserve",
    // invariant: describing a table workbook leaves every input byte untouched
    "formulas-structured-ref": "preserve",

    // ABSENT, deliberately - see UNDECIDED_DESCRIBE_STRUCTURES below.
  },
  /**
   * Unprotect removes only the worksheet (`sheetProtection`) and workbook
   * (`workbookProtection`) protection elements. Neither is a tracked structure,
   * and both are empty attribute-only elements the package layer drops by name,
   * so every tracked structure is carried through untouched: `preserve`.
   *
   * The evidence is `unprotect.corpus.test.ts`, which builds a workbook holding
   * every structure the generator can emit, protects it, unprotects it, and
   * asserts that each output part equals its input part with only the protection
   * element removed. That is a per-cell proof, not an argument: a structure that
   * changed would fail the part comparison. The one absence is external links,
   * left undeclared for the same corpus reason as merge and describe (see
   * UNDECIDED_UNPROTECT_STRUCTURES).
   */
  unprotect: {
    // pins: every output part equals its input with only the protection removed
    "merged-cells": "preserve",
    // pins: every output part equals its input with only the protection removed
    "conditional-formatting": "preserve",
    // pins: every output part equals its input with only the protection removed
    "data-validation": "preserve",
    // pins: every output part equals its input with only the protection removed
    hyperlinks: "preserve",
    // pins: every output part equals its input with only the protection removed
    // (the comment part and its legacy VML drawing travel byte-identical)
    comments: "preserve",
    // pins: every output part equals its input with only the protection removed
    "drawings-charts": "preserve",
    // pins: the workbook part keeps its defined names; only workbookProtection goes
    "defined-names": "preserve",
    // pins: the worksheet keeps its tableParts and the table part is byte-identical
    "excel-tables": "preserve",
    // pins: the totals row is worksheet content and is carried through untouched
    "excel-table-totals-row": "preserve",
    // pins: the pivot table and its cache parts travel byte-identical
    "pivot-tables": "preserve",
    // pins: the calculation chain part is carried through byte-identical
    "calc-chain": "preserve",
    // pins: the shared string store is carried through byte-identical
    "shared-strings": "preserve",
    // pins: the styles part is carried through byte-identical
    "styles-number-formats": "preserve",
    // pins: a macro package keeps its vbaProject.bin and its macro content type
    "vba-project": "preserve",
    // pins: every formula cell is carried through with its cached value untouched
    "formulas-cached": "preserve",
    // pins: an uncached formula is carried through untouched
    "formulas-uncached": "preserve",
    // pins: a shared formula's ref and si are carried through untouched
    "formulas-shared": "preserve",
    // pins: an array formula's ref and text are carried through untouched
    "formulas-array": "preserve",
    // pins: an A1 formula is carried through untouched
    "formulas-a1": "preserve",
    // pins: a structured-reference formula is carried through untouched
    "formulas-structured-ref": "preserve",

    // ABSENT, deliberately - see UNDECIDED_UNPROTECT_STRUCTURES below.
  },
};

/**
 * Why each split cell is still absent. Kept beside the table so the debt is
 * readable without cross-referencing the completeness test.
 */
export const UNDECIDED_SPLIT_STRUCTURES: Readonly<Record<string, string>> = {
  "drawings-charts":
    "No corpus fixture yet: the generator only emits the legacy VML drawing that belongs to a comment.",
  "defined-names":
    "Checked in Phase 1 and deliberately left undeclared: the L1 invariant pass reads defined names but does not relocate them, so the honest answer is still open rather than `fix`. A workbook-scoped name over a filtered region keeps its source rows today.",
  "excel-table-totals-row":
    "The bindings disagree (pins: the range binding deletes the footer block, the table binding keeps it); one cell cannot state both.",
  "external-links":
    "No corpus fixture yet; the generator has no external-link part.",
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

/**
 * Why describe's one cell is absent.
 *
 * The behavior is not in doubt here - a read-only operation cannot disturb an
 * external link any more than it disturbs anything else. The cell stays absent
 * on the corpus rule rather than the decision rule: "a cell the corpus cannot
 * exercise is not a contract", and the generator still has no external-link
 * part to exercise it with. It lands with the fixture, alongside merge's.
 */
export const UNDECIDED_DESCRIBE_STRUCTURES: Readonly<Record<string, string>> = {
  "external-links":
    "No corpus fixture yet; the generator has no external-link part. The answer is certainly `preserve` - describing writes nothing - but an undeclarable cell is left absent rather than asserted from reasoning alone.",
};

/**
 * Why unprotect's one cell is absent.
 *
 * Unprotect changes only the protection elements, so an external link is
 * carried through untouched - its part is byte-identical and its
 * `externalReferences` list, which is not an empty element, is never matched by
 * the strip. The cell stays absent on the corpus rule rather than the decision
 * rule: "a cell the corpus cannot exercise is not a contract", and the
 * generator still has no external-link part to build the fixture from. It lands
 * with that fixture, alongside merge's and describe's.
 */
export const UNDECIDED_UNPROTECT_STRUCTURES: Readonly<Record<string, string>> =
  {
    "external-links":
      "No corpus fixture yet; the generator has no external-link part. Unprotect touches only the protection elements, so a link is certainly carried through untouched, but an undeclarable cell is left absent rather than asserted from reasoning alone.",
  };
