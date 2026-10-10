/**
 * L3: a workbook split to CSV files, every worksheet that carries the column
 * at once (ADR 0007). A CSV file holds one sheet, so each worksheet gives one
 * file per value, each a compact region split of that worksheet.
 */
import {
  ConsultChimpsError,
  throwIfAborted,
  type AbortOutputContext,
  type RandomAccessSource,
} from "@consultchimps/core";

import { XLSX_ERRORS } from "../errors.js";
import { openSheetBook } from "../operations/sheet-book.js";
import { SPLIT_OPERATION } from "./all-worksheet.js";
import { splitOutputFilenames, splitSheetOutputFilenames } from "./names.js";
import {
  resolveRegionSplit,
  type RegionSplitContext,
  type RegionSplitSelection,
  type ResolvedRegionSplit,
} from "./region-split.js";

/** A worksheet a split leaves out, because it gives no rows under the column. */
const SKIPPED = new Set<string>([
  XLSX_ERRORS.XLSX_SPLIT_COLUMN_NOT_FOUND,
  XLSX_ERRORS.XLSX_SPLIT_NO_TABLE,
  XLSX_ERRORS.XLSX_SPLIT_NO_GROUPS,
  // The table reader's own refusal of a column the header row does not name.
  "TABLE_COLUMN_NOT_FOUND",
]);

export interface EachWorksheetSplit {
  /** One region split per worksheet that carries the column, in sheet order. */
  readonly parts: readonly ResolvedRegionSplit[];
  /** One name per group of every part, in part then group order. */
  readonly outputNames: readonly string[];
  /** The worksheets left out because they do not carry the column. */
  readonly skippedSheets: readonly string[];
  /** Rows read in worksheets left out for having no group, all skipped. */
  readonly leftOutRows: LeftOutRows;
}

/** The rows a split read in worksheets it left out, counted in its metrics. */
export interface LeftOutRows {
  readonly inputRows: number;
  readonly skippedRows: number;
}

/** No rows left out, for a split of one region. */
export const NO_LEFT_OUT_ROWS: LeftOutRows = { inputRows: 0, skippedRows: 0 };

/**
 * Resolve a split of every visible worksheet, hidden ones too when
 * `includeHiddenSheets` asks, since each becomes files of its own. One worksheet split keeps the names a
 * one-source split gives, `<prefix>-<value>.csv`; several are named
 * `<prefix>-<value> - <sheet>.csv` so each worksheet's files stay apart.
 */
export async function resolveEachWorksheetSplit(
  source: RandomAccessSource,
  context: RegionSplitContext,
  selection: RegionSplitSelection,
  filenamePrefix: string,
  extension: string,
  /** Checked before each worksheet, so a cancel stops before the next one. */
  abort: {
    signal: AbortSignal | undefined;
    outputContext: AbortOutputContext;
  },
): Promise<EachWorksheetSplit> {
  const book = await openSheetBook(source, {
    file: context.file,
    source: context.label,
    details: context.details,
  });
  const parts: ResolvedRegionSplit[] = [];
  const skippedSheets: string[] = [];
  // Why the sheets were left out, by kind, so a workbook with nothing to split
  // is refused for the reason its sheets gave.
  // A sheet that holds the column but no group is closest to splitting, so its
  // reason wins over the others.
  let columnMissing = false;
  let noGroups: ConsultChimpsError | undefined;
  let lastSkip: ConsultChimpsError | undefined;
  let leftOutInputRows = 0;
  let leftOutSkippedRows = 0;
  for (const sheet of book.sheets) {
    if (sheet.part === undefined) continue;
    // A hidden worksheet's rows leave the workbook only when asked for.
    if (!sheet.visible && selection.includeHiddenSheets !== true) continue;
    throwIfAborted(abort.signal, SPLIT_OPERATION, abort.outputContext);
    try {
      // One open workbook for every sheet, so its parts and strings are held
      // once however many worksheets the split writes.
      parts.push(
        await resolveRegionSplit(
          source,
          context,
          { ...selection, sheet: sheet.name, includeHiddenSheets: true },
          false,
          book,
        ),
      );
    } catch (error) {
      const uncached = (
        error instanceof ConsultChimpsError
          ? (error.details as { uncachedFormulas?: unknown } | undefined)
              ?.uncachedFormulas
          : undefined
      ) as readonly unknown[] | undefined;
      // A sheet that looks empty because Excel never calculated its formulas
      // is refused with that reason rather than left out.
      if (
        error instanceof ConsultChimpsError &&
        SKIPPED.has(error.code) &&
        (uncached === undefined || uncached.length === 0)
      ) {
        skippedSheets.push(sheet.name);
        lastSkip = error;
        if (error.code === XLSX_ERRORS.XLSX_SPLIT_NO_GROUPS) {
          noGroups ??= error;
          // Its rows were read and skipped as blank, so the metrics count them.
          const counts = error.details as
            { inputRows?: number; skippedRows?: number } | undefined;
          leftOutInputRows += counts?.inputRows ?? 0;
          leftOutSkippedRows += counts?.skippedRows ?? 0;
        } else if (error.code !== XLSX_ERRORS.XLSX_SPLIT_NO_TABLE)
          columnMissing = true;
        continue;
      }
      throw error;
    }
  }
  if (parts.length === 0 && noGroups !== undefined) throw noGroups;
  if (parts.length === 0 && !columnMissing && lastSkip !== undefined) {
    throw lastSkip;
  }
  if (parts.length === 0) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_SPLIT_COLUMN_NOT_FOUND,
      `Column "${selection.column}" was not found in any worksheet. Check the header text or provide the correct header row.`,
      {
        details: {
          availableWorksheets: book.sheets.map((sheet) => sheet.name),
          column: selection.column,
          headerRow: selection.headerRow,
          ...context.details,
        },
      },
    );
  }
  const outputNames =
    parts.length === 1
      ? splitOutputFilenames(
          filenamePrefix,
          parts[0]!.groups.map((group) => group.value),
          extension,
        )
      : splitSheetOutputFilenames(
          filenamePrefix,
          parts.flatMap((part) =>
            part.groups.map((group) => ({
              value: group.value,
              sheet: part.sheetName,
            })),
          ),
          extension,
        );
  return {
    parts,
    outputNames,
    skippedSheets,
    leftOutRows: {
      inputRows: leftOutInputRows,
      skippedRows: leftOutSkippedRows,
    },
  };
}

/** The warning naming the worksheets a split left out. */
export function skippedSheetsWarning(
  sheets: readonly string[],
  column: string,
): string[] {
  if (sheets.length === 0) return [];
  return [
    `${sheets.length === 1 ? "Worksheet" : "Worksheets"} ${sheets
      .map((sheet) => `"${sheet}"`)
      .join(
        ", ",
      )} ${sheets.length === 1 ? "has" : "have"} no rows under column "${column}", so no CSV file was written for ${sheets.length === 1 ? "it" : "them"}.`,
  ];
}

/** What a split of one or more region parts reports, summed over the parts. */
export interface RegionSplitSummary {
  readonly groups: ResolvedRegionSplit["groups"];
  /** Distinct values across every part, a value on two worksheets once. */
  readonly distinctGroups: number;
  readonly inputRows: number;
  readonly skippedRows: number;
  readonly uncachedFormulas: readonly string[];
  readonly column: string;
  readonly workbook: ResolvedRegionSplit["workbook"];
  readonly sheets: number;
}

export function regionSplitSummary(
  parts: readonly ResolvedRegionSplit[],
  leftOutRows: LeftOutRows = NO_LEFT_OUT_ROWS,
): RegionSplitSummary {
  const first = parts[0]!;
  let inputRows = leftOutRows.inputRows;
  let skippedRows = leftOutRows.skippedRows;
  const groups: ResolvedRegionSplit["groups"][number][] = [];
  const uncachedFormulas: string[] = [];
  for (const part of parts) {
    inputRows += part.inputRows;
    skippedRows += part.skippedRows;
    for (const group of part.groups) groups.push(group);
    for (const location of part.uncachedFormulas)
      uncachedFormulas.push(location);
  }
  return {
    groups,
    distinctGroups: new Set(groups.map((group) => group.key)).size,
    inputRows,
    skippedRows,
    uncachedFormulas,
    column: first.column,
    workbook: first.workbook,
    sheets: parts.length,
  };
}
