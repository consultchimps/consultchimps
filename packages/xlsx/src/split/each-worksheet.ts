/**
 * L3: a workbook split to CSV files, every worksheet that carries the column
 * at once (ADR 0007). A CSV file holds one sheet, so each worksheet gives one
 * file per value, each a compact region split of that worksheet.
 */
import {
  ConsultChimpsError,
  type RandomAccessSource,
} from "@consultchimps/core";

import { XLSX_ERRORS } from "../errors.js";
import { openSheetBook } from "../operations/sheet-book.js";
import { splitOutputFilenames, splitSheetOutputFilenames } from "./names.js";
import {
  resolveRegionSplit,
  type RegionSplitContext,
  type RegionSplitSelection,
  type ResolvedRegionSplit,
} from "./region-split.js";

/** A worksheet a split leaves out, because it has no rows under the column. */
const SKIPPED = new Set<string>([
  XLSX_ERRORS.XLSX_SPLIT_COLUMN_NOT_FOUND,
  XLSX_ERRORS.XLSX_SPLIT_NO_TABLE,
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
}

/**
 * Resolve a split of every worksheet, hidden ones included as the
 * whole-workbook split includes them. One worksheet split keeps the names a
 * one-source split gives, `<prefix>-<value>.csv`; several are named
 * `<prefix>-<value> - <sheet>.csv` so each worksheet's files stay apart.
 */
export async function resolveEachWorksheetSplit(
  source: RandomAccessSource,
  context: RegionSplitContext,
  selection: RegionSplitSelection,
  filenamePrefix: string,
  extension: string,
): Promise<EachWorksheetSplit> {
  const book = await openSheetBook(source, {
    file: context.file,
    source: context.label,
    details: context.details,
  });
  const parts: ResolvedRegionSplit[] = [];
  const skippedSheets: string[] = [];
  for (const sheet of book.sheets) {
    if (sheet.part === undefined) continue;
    try {
      parts.push(
        await resolveRegionSplit(
          source,
          context,
          { ...selection, sheet: sheet.name, includeHiddenSheets: true },
          false,
        ),
      );
    } catch (error) {
      if (error instanceof ConsultChimpsError && SKIPPED.has(error.code)) {
        skippedSheets.push(sheet.name);
        continue;
      }
      throw error;
    }
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
  return { parts, outputNames, skippedSheets };
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
  readonly inputRows: number;
  readonly skippedRows: number;
  readonly uncachedFormulas: readonly string[];
  readonly column: string;
  readonly workbook: ResolvedRegionSplit["workbook"];
  readonly sheets: number;
}

export function regionSplitSummary(
  parts: readonly ResolvedRegionSplit[],
): RegionSplitSummary {
  const first = parts[0]!;
  let inputRows = 0;
  let skippedRows = 0;
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
    inputRows,
    skippedRows,
    uncachedFormulas,
    column: first.column,
    workbook: first.workbook,
    sheets: parts.length,
  };
}
