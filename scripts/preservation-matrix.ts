/**
 * The documentation projection of the xlsx conformance contract.
 *
 * `@consultchimps/xlsx` exports the contract (`packages/xlsx/src/contract.ts`)
 * as the single source of truth for what each Excel operation promises to do
 * to each workbook structure. It is written for the package's own tests, in
 * the package's vocabulary. This module is the one place that translates it
 * for readers of the documentation site: a human label and one short note per
 * structure, a plain-language status per declared behavior, and the markdown
 * block that `apps/docs/content/docs/tools/excel-preservation.mdx` carries.
 *
 * It is read through the package's public entry point rather than its source
 * tree, so this projection sees exactly the surface a consumer sees, and
 * `pnpm docs:check` needs a build first - as the CLI reference check needs the
 * built CLI.
 *
 * Nothing here may invent a promise. Every status in the rendered table is
 * derived from a contract cell, and a structure the contract deliberately
 * leaves undeclared renders as "Needs review" rather than as a claim the
 * conformance tests do not hold up. `collectProjectionProblems` fails
 * `pnpm docs:check` when the contract grows a structure, an operation, or a
 * behavior this module has no words for, so the published matrix cannot go
 * stale quietly.
 */
import {
  CONTRACT,
  OPERATIONS,
  TRACKED_STRUCTURES,
  UNDECIDED_DESCRIBE_STRUCTURES,
  UNDECIDED_MERGE_STRUCTURES,
  UNDECIDED_SPLIT_STRUCTURES,
  UNDECIDED_UNPROTECT_STRUCTURES,
  type ContractBehavior,
  type ContractOperation as Operation,
  type ContractStructure as Structure,
} from "@consultchimps/xlsx";

/** How the site names one workbook structure, and what a reader should know. */
export interface StructureDocumentation {
  /** Column-one label. Plain English, not the contract's key. */
  readonly label: string;
  /** One line in the "What to check" column. */
  readonly note: string;
}

/** How the site names one operation. */
export interface OperationDocumentation {
  /** Table column heading, or the bold lead-in of an aside entry. */
  readonly label: string;
  /**
   * Why an operation has no column. Rendered only for an operation the
   * contract declares no cell for; kept for every operation so a column that
   * later disappears cannot leave the page unexplained.
   */
  readonly undeclaredNote: string;
}

/** One cell status and the legend row that explains it. */
export interface StatusDocumentation {
  readonly label: string;
  readonly meaning: string;
}

export const GENERATED_BLOCK_START = "{/* preservation-matrix:start */}";
export const GENERATED_BLOCK_END = "{/* preservation-matrix:end */}";

/**
 * A note is a single table cell, so it stays a sentence or two. Anything
 * longer belongs in the hand-written prose around the generated block.
 */
const MAX_NOTE_LENGTH = 220;

/** A qualifier shares a cell with its status, so it stays a clause. */
const MAX_QUALIFIER_LENGTH = 120;

export const STRUCTURE_DOCUMENTATION: Record<
  Structure,
  StructureDocumentation
> = {
  "merged-cells": {
    label: "Merged cells",
    note: "Merged ranges follow the rows they cover, so no merge is left spanning rows an output does not contain.",
  },
  "conditional-formatting": {
    label: "Conditional formatting",
    note: "Rules travel with the worksheet together with the formats they name, and a split shrinks the range a rule covers.",
  },
  "data-validation": {
    label: "Data validation",
    note: "Dropdown lists and entry rules keep their cells, and a split shrinks the range they apply to.",
  },
  hyperlinks: {
    label: "Hyperlinks",
    note: "A link stays on the cell it decorates and keeps its target.",
  },
  comments: {
    label: "Cell comments",
    note: "A comment and the box that draws it follow the row they annotate.",
  },
  "drawings-charts": {
    label: "Images, shapes, and charts",
    note: "Pictures and charts travel with their worksheet. A merge rewrites a chart's references when a rename forces it; a split never re-points them, so open each chart in a split output and check what it plots.",
  },
  "defined-names": {
    label: "Defined names",
    note: "A merge suffixes a name a later input also claimed but does not update formulas that used it. A split does not move a name's coordinates, so one over filtered rows still spans its original rows.",
  },
  "excel-tables": {
    label: "Excel Tables",
    note: "Table parts travel with their worksheet: a split resizes the table and its filter, and a merge renames a table whose name another input already claimed.",
  },
  "excel-table-totals-row": {
    label: "Excel Table totals rows",
    note: "The default whole-workbook split keeps the totals row; the compact single-source modes rebuild data only and drop it. Check the foot of each table.",
  },
  "pivot-tables": {
    label: "Pivot tables and their caches",
    note: "A pivot cache is a private copy of its source rows, so a split and a merge both remove the pivot and its cache and say so. Rebuild it in Excel from the output's own rows.",
  },
  "calc-chain": {
    label: "Calculation chain",
    note: "Excel's internal recalculation index. No output keeps an entry for a row that is gone, and a merged workbook asks Excel to rebuild the index when it opens.",
  },
  "shared-strings": {
    label: "Shared text store",
    note: "The workbook's internal table of cell text. A merge keeps one store and re-points every copied cell at it, so no text is lost or duplicated.",
  },
  "styles-number-formats": {
    label: "Cell styles and number formats",
    note: "Fonts, fills, borders, alignment, and number formats stay on their cells; a merge copies and de-duplicates the ones its worksheets use.",
  },
  "vba-project": {
    label: "Macros (VBA project)",
    note: "A split keeps macros and writes .xlsm. A merge carries them only when the first input is the one that has them, no other input does, and the output is named .xlsm; otherwise they are removed and reported.",
  },
  "external-links": {
    label: "Links to other workbooks",
    note: "A split carries the link but never re-points it, and a merge cannot interleave two workbooks' link lists, so it removes them and reports it. Check every link before delivery.",
  },
  "formulas-cached": {
    label: "Formulas with saved results",
    note: "The formula travels and its references follow the rows that moved. The result Excel last saved travels with it and is not recalculated, so check any total taken over rows an output no longer holds.",
  },
  "formulas-uncached": {
    label: "Formulas Excel never calculated",
    note: "The formula is carried through unchanged. In values-only mode it has no saved result to bake, so the cell becomes a formatted blank and the result names the worksheet and cell.",
  },
  "formulas-shared": {
    label: "Shared formulas filled across a range",
    note: "The range a shared formula spans shrinks with the rows it covers, so the fill stays valid.",
  },
  "formulas-array": {
    label: "Array formulas",
    note: "The array's span and its references follow the rows. One array formula anywhere on a sheet stops that sheet's Excel Table from being compacted, which is reported.",
  },
  "formulas-a1": {
    label: "Formulas using ordinary cell references",
    note: "References such as =SUM(B2:B40) are rewritten when the rows they name move, and a merge follows a worksheet that a name collision renamed.",
  },
  "formulas-structured-ref": {
    label: "Formulas using Excel Table names",
    note: "A structured reference names a table column rather than a cell address, so it survives a split unchanged and follows a table a merge had to rename.",
  },
};

export const OPERATION_DOCUMENTATION: Record<
  Operation,
  OperationDocumentation
> = {
  split: {
    label: "Split by column",
    undeclaredNote:
      "The workbook splitter makes no checked promise about any structure yet.",
  },
  merge: {
    label: "Merge tabs",
    undeclaredNote:
      "The workbook merge makes no checked promise about any structure yet.",
  },
  consolidate: {
    label: "Consolidate rows",
    undeclaredNote:
      "Consolidation reads stored cell values and writes a new single-worksheet table, so nothing from the source package travels with the rows: no formatting, no formulas, no tables, no macros. A table-aware consolidation that could carry more is planned, and the contract stays silent until it ships.",
  },
  values: {
    label: "Values-only output",
    undeclaredNote:
      "Values-only is an option on the operations above rather than an operation of its own. It replaces each formula with the result the source file last saved and leaves every other structure to the operation it runs on, so the rest of this table still applies.",
  },
  describe: {
    label: "Inspect",
    undeclaredNote:
      "The workbook inspection makes no checked promise about any structure yet.",
  },
  unprotect: {
    label: "Unprotect",
    undeclaredNote:
      "The workbook unprotect makes no checked promise about any structure yet.",
  },
};

/**
 * The site's plain-language name for each contract behavior. The legend is
 * rendered from the statuses the table actually uses, so a behavior no cell
 * declares cannot introduce a legend row nobody can find in the matrix.
 */
export const BEHAVIOR_STATUS: Record<ContractBehavior, StatusDocumentation> = {
  preserve: {
    label: "Preserved",
    meaning:
      "Carried into the output untouched, and still correct there. Nothing to check.",
  },
  fix: {
    label: "Adjusted to stay correct",
    meaning:
      "Rewritten as part of the operation so that it still points at the right rows, cells, or names once the output is open in Excel.",
  },
  "strip-warn": {
    label: "Removed, reported as a warning",
    meaning:
      "Deliberately left out of the output, and the operation's result says so in plain language rather than removing it quietly.",
  },
  refuse: {
    label: "Refused before anything is written",
    meaning:
      "The operation stops with an explanatory error instead of writing an output that would be wrong.",
  },
};

/**
 * The status for a structure the contract deliberately leaves undeclared. The
 * absent cells are the package's own debt ledger, so this status promises
 * nothing at all: what the operation does today is stated in the row's note and
 * in its qualifier, not implied by the status.
 */
export const NEEDS_REVIEW_STATUS: StatusDocumentation = {
  label: "Needs review",
  meaning:
    "Not covered by a checked promise, so the package may still change what it does here. The note on the row says what happens today; open the output and check it before sending the file on.",
};

/** What a cell of the contract holds: a declared behavior, or nothing yet. */
export type CellKind = ContractBehavior | "undeclared";

/**
 * Statuses that replace the shared wording for one operation, because the
 * shared wording describes an output that operation never produces. The
 * inspection is the case: it writes nothing, so "carried into the output" is a
 * promise it cannot make and "check the output before sending it on" is advice
 * about a file that does not exist. What it does promise - your file is
 * untouched - its corpus proves by comparing the input's bytes before and after.
 */
export const OPERATION_STATUS_OVERRIDES: Partial<
  Record<Operation, Partial<Record<CellKind, StatusDocumentation>>>
> = {
  describe: {
    preserve: {
      label: "Read only, nothing touched",
      meaning:
        "The inspection produces no file. It reads the workbook, reports what it found, and leaves every byte of your original in place.",
    },
    undeclared: {
      label: "Read only, not yet promised",
      meaning:
        "The inspection has no recorded promise about this structure yet. It still writes nothing, so nothing in your file can change; the note on the row says what it does with the structure today.",
    },
  },
};

/**
 * Per-cell qualifiers, for the cells where one contract word is coarser than
 * the behavior a reader has to plan around: a promise that holds only under a
 * condition, or an undeclared cell whose shipped behavior is not the neutral
 * "carried as it stands". The status still comes from the contract; the
 * qualifier only says what the single word leaves out, and is checked to name
 * a tracked structure in an operation that has a column.
 */
export const CELL_QUALIFIERS: Partial<
  Record<Operation, Partial<Record<Structure, string>>>
> = {
  split: {
    // The cell is `preserve` and true of the formula: its references follow
    // the rows. The saved result is another matter - a plain split leaves it
    // as Excel wrote it, and only a values-only split clears one computed over
    // removed rows (src/tier1/stale-values.ts), so "nothing to check" would be
    // the wrong thing to tell someone reading numbers before a recalculation.
    "formulas-cached":
      "the saved result is not recalculated, so a total over removed rows keeps its old answer",
  },
  merge: {
    // The contract cell is `strip-warn`, but the removal is conditional: the
    // project travels when the seed - the first input - is the only one
    // carrying it and the output name admits macros (`resolveMacroProject`
    // requires all three). See ARCHITECTURE.md, "The merge's removals".
    "vba-project":
      "kept when the first input is the only one with macros and the output is named .xlsm",
    // The cell is absent because no corpus fixture can exercise it yet, but
    // the transplant already removes external links and warns about it, so
    // the bare status would read as more preservation than the merge offers.
    "external-links":
      "the merge removes them and says so; the contract cannot pin that until a fixture exists",
  },
  unprotect: {
    // The cell is absent for the same corpus reason, but unprotect touches only
    // the protection elements, so a link is carried through untouched; the bare
    // "Needs review" would read as less preservation than unprotect offers.
    "external-links":
      "unprotect changes only the protection, so a link travels unchanged; no fixture pins that yet",
  },
};

/**
 * Where each operation records the structures it has not decided yet. An
 * operation that declares a cell must appear here, so a new column cannot
 * silently render a status with no recorded reason behind it.
 */
const UNDECIDED_STRUCTURES_BY_OPERATION: Partial<
  Record<Operation, Readonly<Record<string, string>>>
> = {
  split: UNDECIDED_SPLIT_STRUCTURES,
  merge: UNDECIDED_MERGE_STRUCTURES,
  describe: UNDECIDED_DESCRIBE_STRUCTURES,
  unprotect: UNDECIDED_UNPROTECT_STRUCTURES,
};

/** Legend order: declared behaviors first, then the undeclared status. */
function statusOrder(): readonly StatusDocumentation[] {
  const ordered: StatusDocumentation[] = [
    BEHAVIOR_STATUS.preserve,
    BEHAVIOR_STATUS.fix,
    BEHAVIOR_STATUS["strip-warn"],
    BEHAVIOR_STATUS.refuse,
  ];
  for (const overrides of Object.values(OPERATION_STATUS_OVERRIDES)) {
    ordered.push(...Object.values(overrides ?? {}));
  }
  ordered.push(NEEDS_REVIEW_STATUS);
  return ordered;
}

function declaredStructureCount(operation: Operation): number {
  return Object.keys(CONTRACT[operation]).length;
}

/** Operations with at least one declared cell; these earn a table column. */
export function columnOperations(): readonly Operation[] {
  return OPERATIONS.filter(
    (operation) => declaredStructureCount(operation) > 0,
  );
}

/** Operations the contract declares nothing for; these are explained in prose. */
export function undeclaredOperations(): readonly Operation[] {
  return OPERATIONS.filter(
    (operation) => declaredStructureCount(operation) === 0,
  );
}

/**
 * The status the contract implies for one cell of the published matrix. An
 * absent cell is the contract's own debt ledger, so it renders as "Needs
 * review"; a behavior with no status cannot reach here, because
 * `collectProjectionProblems` fails the check before anything is rendered.
 */
export function statusFor(
  operation: Operation,
  structure: Structure,
): StatusDocumentation {
  const behavior = CONTRACT[operation][structure];
  const override =
    OPERATION_STATUS_OVERRIDES[operation]?.[behavior ?? "undeclared"];
  if (override !== undefined) {
    return override;
  }
  return behavior === undefined
    ? NEEDS_REVIEW_STATUS
    : BEHAVIOR_STATUS[behavior];
}

/** The rendered cell: the contract's status, plus a qualifier where one exists. */
export function cellTextFor(
  operation: Operation,
  structure: Structure,
): string {
  const status = statusFor(operation, structure);
  const qualifier = CELL_QUALIFIERS[operation]?.[structure];
  return qualifier === undefined
    ? status.label
    : `${status.label} (${qualifier})`;
}

/**
 * Everything the projection cannot express about the contract as it stands.
 * An empty list is the only acceptable result; the check script turns anything
 * else into a failure that names the missing label, note, or status.
 */
export function collectProjectionProblems(): string[] {
  const problems: string[] = [];
  const trackedStructures = new Set<string>(TRACKED_STRUCTURES);

  for (const structure of TRACKED_STRUCTURES) {
    const documentation = STRUCTURE_DOCUMENTATION[structure] as
      StructureDocumentation | undefined;
    if (documentation === undefined) {
      problems.push(
        `the contract tracks the structure "${structure}", which has no human label or note in STRUCTURE_DOCUMENTATION`,
      );
      continue;
    }
    if (documentation.label.trim() === "") {
      problems.push(`structure "${structure}" has an empty label`);
    }
    if (documentation.note.trim() === "") {
      problems.push(`structure "${structure}" has an empty note`);
    }
    if (documentation.note.length > MAX_NOTE_LENGTH) {
      problems.push(
        `structure "${structure}" has a ${documentation.note.length}-character note; keep it under ${MAX_NOTE_LENGTH} characters and move the detail into the page's prose`,
      );
    }
  }

  for (const structure of Object.keys(STRUCTURE_DOCUMENTATION)) {
    if (!trackedStructures.has(structure)) {
      problems.push(
        `STRUCTURE_DOCUMENTATION describes "${structure}", which the contract no longer tracks`,
      );
    }
  }

  for (const operation of OPERATIONS) {
    const documentation = OPERATION_DOCUMENTATION[operation] as
      OperationDocumentation | undefined;
    if (documentation === undefined) {
      problems.push(
        `the contract declares the operation "${operation}", which has no human label in OPERATION_DOCUMENTATION`,
      );
      continue;
    }
    if (documentation.label.trim() === "") {
      problems.push(`operation "${operation}" has an empty label`);
    }
    if (documentation.undeclaredNote.trim() === "") {
      problems.push(
        `operation "${operation}" has no note explaining an absent column`,
      );
    }
  }

  const declaredOperations = new Set<string>(OPERATIONS);
  for (const operation of Object.keys(OPERATION_DOCUMENTATION)) {
    if (!declaredOperations.has(operation)) {
      problems.push(
        `OPERATION_DOCUMENTATION describes "${operation}", which is not an operation of the contract`,
      );
    }
  }

  for (const operation of columnOperations()) {
    const undecided = UNDECIDED_STRUCTURES_BY_OPERATION[operation];
    if (undecided === undefined) {
      problems.push(
        `operation "${operation}" declares contract cells but scripts/preservation-matrix.ts does not know where its undecided structures are recorded; wire its UNDECIDED_* record into UNDECIDED_STRUCTURES_BY_OPERATION`,
      );
      continue;
    }
    for (const structure of TRACKED_STRUCTURES) {
      if (
        CONTRACT[operation][structure] === undefined &&
        undecided[structure] === undefined
      ) {
        problems.push(
          `the contract declares no behavior for ${operation}.${structure} and records no reason for it, so the published matrix would show "${NEEDS_REVIEW_STATUS.label}" with nothing behind it`,
        );
      }
    }
  }

  for (const operation of OPERATIONS) {
    for (const [structure, behavior] of Object.entries(CONTRACT[operation])) {
      if (BEHAVIOR_STATUS[behavior] === undefined) {
        problems.push(
          `${operation}.${structure} declares the behavior "${behavior}", which has no plain-language status in BEHAVIOR_STATUS`,
        );
      }
    }
  }

  const operationsWithColumn = new Set<string>(columnOperations());
  for (const [operation, overrides] of Object.entries(
    OPERATION_STATUS_OVERRIDES,
  )) {
    if (!operationsWithColumn.has(operation)) {
      problems.push(
        `OPERATION_STATUS_OVERRIDES gives "${operation}" its own statuses, but it has no column in the matrix, so they would never be rendered`,
      );
      continue;
    }
    for (const [kind, status] of Object.entries(overrides ?? {})) {
      if (
        kind !== "undeclared" &&
        BEHAVIOR_STATUS[kind as ContractBehavior] === undefined
      ) {
        problems.push(
          `OPERATION_STATUS_OVERRIDES overrides ${operation}.${kind}, which is neither a contract behavior nor "undeclared"`,
        );
        continue;
      }
      if (status.label.trim() === "" || status.meaning.trim() === "") {
        problems.push(
          `the ${operation} status for "${kind}" needs both a label and a meaning`,
        );
      }
    }
  }

  // Two statuses sharing a label would collapse into one legend row, and a
  // reader would have no way to tell which of them a cell meant.
  const meaningByLabel = new Map<string, string>();
  for (const status of statusOrder()) {
    const existing = meaningByLabel.get(status.label);
    if (existing !== undefined && existing !== status.meaning) {
      problems.push(
        `two statuses are both labelled "${status.label}" with different meanings; give one of them its own wording`,
      );
    }
    meaningByLabel.set(status.label, status.meaning);
  }

  for (const [operation, qualifiers] of Object.entries(CELL_QUALIFIERS)) {
    if (!operationsWithColumn.has(operation)) {
      problems.push(
        `CELL_QUALIFIERS qualifies cells of "${operation}", which has no column in the matrix, so the qualifier would never be rendered`,
      );
      continue;
    }
    for (const [structure, qualifier] of Object.entries(qualifiers ?? {})) {
      if (!trackedStructures.has(structure)) {
        problems.push(
          `CELL_QUALIFIERS qualifies ${operation}.${structure}, which the contract does not track`,
        );
        continue;
      }
      if (qualifier.trim() === "") {
        problems.push(`the qualifier on ${operation}.${structure} is empty`);
      }
      if (qualifier.length > MAX_QUALIFIER_LENGTH) {
        problems.push(
          `the qualifier on ${operation}.${structure} is ${qualifier.length} characters; keep it under ${MAX_QUALIFIER_LENGTH} so the cell stays readable`,
        );
      }
    }
  }

  return problems;
}

/** One markdown table. Cell padding is left to Prettier. */
function renderTable(
  headings: readonly string[],
  rows: readonly (readonly string[])[],
): string {
  const line = (cells: readonly string[]): string => `| ${cells.join(" | ")} |`;
  return [
    line(headings),
    line(headings.map(() => "---")),
    ...rows.map(line),
  ].join("\n");
}

/** The legend, showing only the statuses the matrix below it actually uses. */
function renderStatusLegend(usedStatuses: ReadonlySet<string>): string {
  return renderTable(
    ["Status", "What it means"],
    statusOrder()
      .filter((status) => usedStatuses.has(status.label))
      .map((status) => [status.label, status.meaning]),
  );
}

/**
 * The generated block: the status legend, the structure by operation matrix,
 * and an entry for every operation the contract declares nothing for.
 */
export function renderPreservationMatrixBlock(): string {
  const operations = columnOperations();
  const usedStatuses = new Set<string>();
  const rows = TRACKED_STRUCTURES.map((structure) => {
    const documentation = STRUCTURE_DOCUMENTATION[structure];
    const cells = operations.map((operation) => {
      usedStatuses.add(statusFor(operation, structure).label);
      return cellTextFor(operation, structure);
    });
    return [documentation.label, ...cells, documentation.note];
  });

  const sections = [
    GENERATED_BLOCK_START,
    "## What each status means",
    renderStatusLegend(usedStatuses),
    "## Structure by operation",
    renderTable(
      [
        "Workbook structure",
        ...operations.map(
          (operation) => OPERATION_DOCUMENTATION[operation].label,
        ),
        "What to check",
      ],
      rows,
    ),
  ];

  const withoutColumn = undeclaredOperations();
  if (withoutColumn.length > 0) {
    sections.push(
      "## Operations without a column",
      withoutColumn
        .map((operation) => {
          const documentation = OPERATION_DOCUMENTATION[operation];
          return `- **${documentation.label}.** ${documentation.undeclaredNote}`;
        })
        .join("\n"),
    );
  }

  sections.push(GENERATED_BLOCK_END);
  return `${sections.join("\n\n")}\n`;
}
