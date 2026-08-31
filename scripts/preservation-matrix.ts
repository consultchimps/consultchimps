/**
 * The documentation projection of the xlsx conformance contract.
 *
 * `packages/xlsx/src/contract.ts` is the single source of truth for what each
 * Excel operation promises to do to each workbook structure. It is written for
 * the package's own tests, in the package's vocabulary. This module is the one
 * place that translates it for readers of the documentation site: a human label
 * and one short note per structure, a plain-language status per declared
 * behavior, and the markdown block that
 * `apps/docs/content/docs/tools/excel-preservation.mdx` carries.
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
  type ContractBehavior,
  type Operation,
  type Structure,
} from "../packages/xlsx/src/contract.ts";

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
    note: "Pictures and charts are carried with their worksheet, but a chart's source range is never re-pointed: open each chart in an output and check what it plots.",
  },
  "defined-names": {
    label: "Defined names",
    note: "A merge makes a duplicate name unique and updates the formulas that used it. A split does not move a name's coordinates, so a name over filtered rows still spans its original rows.",
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
    note: "A split keeps macros and writes .xlsm. A merge carries them only when exactly one input has them and the output is named .xlsm; otherwise they are removed and reported.",
  },
  "external-links": {
    label: "Links to other workbooks",
    note: "A split carries the link but never re-points it, and a merge cannot interleave two workbooks' link lists, so it removes them and reports it. Check every link before delivery.",
  },
  "formulas-cached": {
    label: "Formulas with saved results",
    note: "The formula and the result Excel last saved both travel unchanged; values-only mode keeps the saved result and drops the formula.",
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
 * absent cells are the package's own debt ledger, so the honest reading is
 * "carried as far as it can be, promised nowhere" rather than a guarantee.
 */
export const NEEDS_REVIEW_STATUS: StatusDocumentation = {
  label: "Needs review",
  meaning:
    "Not yet covered by a checked promise. Whatever the file holds is carried through as it stands, but nothing re-points or rebuilds it, so open the output and check this before sending it on.",
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
};

/** Legend order: declared behaviors first, then the undeclared status. */
const STATUS_ORDER: readonly StatusDocumentation[] = [
  BEHAVIOR_STATUS.preserve,
  BEHAVIOR_STATUS.fix,
  BEHAVIOR_STATUS["strip-warn"],
  BEHAVIOR_STATUS.refuse,
  NEEDS_REVIEW_STATUS,
];

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

/** The status the contract implies for one cell of the published matrix. */
export function statusFor(
  operation: Operation,
  structure: Structure,
): StatusDocumentation {
  const behavior = CONTRACT[operation][structure];
  return behavior === undefined
    ? NEEDS_REVIEW_STATUS
    : (BEHAVIOR_STATUS[behavior] ?? NEEDS_REVIEW_STATUS);
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
    STATUS_ORDER.filter((status) => usedStatuses.has(status.label)).map(
      (status) => [status.label, status.meaning],
    ),
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
      const status = statusFor(operation, structure);
      usedStatuses.add(status.label);
      return status.label;
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
