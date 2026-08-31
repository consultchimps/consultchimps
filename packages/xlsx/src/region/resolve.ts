/**
 * L2 region layer: the one resolver.
 *
 * Every discovery heuristic in the package lands here: NFKC/trim/case-folded
 * header search, the `headerRow` override, and the rule that associates a
 * detected header with an Excel Table when the header sits on that table's
 * header row. The heuristics are ports of `findSplitHeader` and
 * `findMatchingTable` from `src/workbook-column-split.ts`, re-expressed over
 * `WorkbookModel` / `WorksheetModel` instead of SheetJS objects, with row
 * numbers one-based throughout (the old code mixed zero-based indices with
 * one-based row numbers).
 *
 * Search order, and therefore the tiebreak, is inherited unchanged: sheets in
 * workbook order; within a sheet, rows top to bottom; within a row, columns
 * left to right. `{ find }` returns the first sheet that matches under that
 * order, `"all-worksheets"` returns every sheet that matches.
 */

import { ConsultChimpsError } from "@consultchimps/core";

import { XLSX_ERRORS } from "../errors.js";
import type {
  RowNumber,
  WorkbookModel,
  WorkbookTableInfo,
  WorksheetModel,
} from "../model/types.js";
import { RangeBinding } from "./range-binding.js";
import { TableBinding } from "./table-binding.js";
import type {
  DataRegion,
  RegionOrigin,
  RegionSelector,
  ResolveRegions,
} from "./types.js";
import { normalizeHeader, parseSheetRange } from "./values.js";

/** Excel's own reserved names (`_xlnm.Print_Area`, ...) are never regions. */
const BUILTIN_DEFINED_NAME_PREFIX = "_xlnm.";

interface HeaderCell {
  readonly column: number;
  readonly row: RowNumber;
}

/**
 * Port of `findSplitHeader`. With `configuredHeaderRow` only that row is
 * examined, so a header row override never silently matches a data row that
 * happens to repeat the header text.
 */
function findHeaderCell(
  worksheet: WorksheetModel,
  headerText: string,
  configuredHeaderRow: RowNumber | undefined,
): HeaderCell | undefined {
  const used = worksheet.usedRange;
  if (!used) {
    return undefined;
  }
  const target = normalizeHeader(headerText);
  const firstRow = configuredHeaderRow ?? used.start.row;
  const lastRow = configuredHeaderRow ?? used.end.row;
  for (let row = firstRow; row <= lastRow; row += 1) {
    for (
      let column = used.start.column;
      column <= used.end.column;
      column += 1
    ) {
      const text = worksheet.cellText({ column, row });
      if (text !== undefined && normalizeHeader(text) === target) {
        return { column, row };
      }
    }
  }
  return undefined;
}

/**
 * Port of `findMatchingTable`. A detected header belongs to a table when the
 * table is on the same sheet, has a header row, starts on exactly that row,
 * and names the same column at the matching offset.
 *
 * `WorkbookTableInfo` has no "this table has no header row" flag (the old
 * `ExcelTableDefinition.headerRow` was a boolean), so a header row below 1 is
 * read as "headerless" and never associates.
 */
function findMatchingTable(
  tables: readonly WorkbookTableInfo[],
  sheetName: string,
  headerRow: RowNumber,
  columnIndex: number,
  columnText: string,
): WorkbookTableInfo | undefined {
  const target = normalizeHeader(columnText);
  return tables.find((table) => {
    if (table.sheetName !== sheetName || table.headerRow < 1) {
      return false;
    }
    const offset = columnIndex - table.range.start.column;
    return (
      table.range.start.row === headerRow &&
      table.headerRow === headerRow &&
      offset >= 0 &&
      offset < table.columnNames.length &&
      normalizeHeader(table.columnNames[offset] ?? "") === target
    );
  });
}

/**
 * A region for a header that was located on a sheet: an Excel Table when the
 * header belongs to one, a worksheet range down to the last used row
 * otherwise.
 */
function bindingForHeader(
  worksheet: WorksheetModel,
  header: HeaderCell,
  columnText: string,
  tables: readonly WorkbookTableInfo[],
  origin: RegionOrigin,
): DataRegion {
  const table = findMatchingTable(
    tables,
    worksheet.info.name,
    header.row,
    header.column,
    columnText,
  );
  if (table) {
    return new TableBinding(worksheet, table);
  }
  const used = worksheet.usedRange;
  return new RangeBinding({
    body: {
      end: {
        column: used?.end.column ?? header.column,
        row: used?.end.row ?? header.row,
      },
      start: {
        column: used?.start.column ?? header.column,
        row: header.row + 1,
      },
    },
    headerRow: header.row,
    origin,
    worksheet,
  });
}

function requireColumn(
  column: string | undefined,
  selectorLabel: string,
): string {
  if (column === undefined || column.trim() === "") {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_NO_COLUMNS,
      `The ${selectorLabel} selector needs a column name to search for. Provide the header text of the column you want to use.`,
      { details: { selector: selectorLabel } },
    );
  }
  return column;
}

function requireWorksheet(
  workbook: WorkbookModel,
  sheetName: string,
): WorksheetModel {
  const worksheet = workbook.worksheet(sheetName);
  if (!worksheet) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_WORKSHEET_NOT_FOUND,
      `Worksheet "${sheetName}" was not found in the workbook. Check the sheet name and try again.`,
      {
        details: {
          availableWorksheets: workbook.sheets.map((sheet) => sheet.name),
          sheet: sheetName,
        },
      },
    );
  }
  return worksheet;
}

function columnNotFound(
  workbook: WorkbookModel,
  columnText: string,
  headerRow: RowNumber | undefined,
): ConsultChimpsError {
  return new ConsultChimpsError(
    XLSX_ERRORS.XLSX_SPLIT_COLUMN_NOT_FOUND,
    `Column "${columnText}" was not found in any worksheet. Check the header text or provide the correct header row.`,
    {
      details: {
        availableWorksheets: workbook.sheets.map((sheet) => sheet.name),
        column: columnText,
        headerRow,
      },
    },
  );
}

async function resolveTable(
  workbook: WorkbookModel,
  tableName: string,
): Promise<readonly DataRegion[]> {
  const tables = await workbook.tables();
  const target = tableName.toLocaleLowerCase();
  const table = tables.find(
    (candidate) => candidate.name.toLocaleLowerCase() === target,
  );
  if (!table) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_SPLIT_NO_TABLE,
      `Excel Table "${tableName}" was not found in the workbook. Check the table name in Excel under Table Design.`,
      {
        details: {
          availableTables: tables.map((candidate) => candidate.name),
          table: tableName,
        },
      },
    );
  }
  return [new TableBinding(requireWorksheet(workbook, table.sheetName), table)];
}

async function resolveRange(
  workbook: WorkbookModel,
  reference: string,
): Promise<readonly DataRegion[]> {
  const definedNames = await workbook.definedNames();
  const target = reference.trim().toLocaleLowerCase();
  const definedName = definedNames.find(
    (entry) =>
      !entry.name.startsWith(BUILTIN_DEFINED_NAME_PREFIX) &&
      entry.name.toLocaleLowerCase() === target,
  );
  const parsed = parseSheetRange(definedName?.reference ?? reference);
  if (!parsed) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_INVALID_NAMED_RANGE,
      definedName
        ? `Defined name "${reference}" does not point at a single worksheet range.`
        : `"${reference}" is not a defined name in this workbook or a range reference such as "Sheet1!A1:F200".`,
      {
        details: {
          range: definedName?.reference ?? reference,
          resolvedFrom: definedName ? "defined-name" : "reference",
        },
      },
    );
  }
  const worksheet = requireWorksheet(workbook, parsed.sheet);
  return [
    new RangeBinding({
      body: {
        end: parsed.range.end,
        start: {
          column: parsed.range.start.column,
          row: parsed.range.start.row + 1,
        },
      },
      headerRow: parsed.range.start.row,
      origin: definedName
        ? { kind: "named-range", rangeName: definedName.name }
        : { kind: "explicit-range", reference },
      worksheet,
    }),
  ];
}

async function resolveSheet(
  workbook: WorkbookModel,
  sheetName: string,
  headerRow: RowNumber | undefined,
  column: string | undefined,
): Promise<readonly DataRegion[]> {
  if (
    headerRow !== undefined &&
    (!Number.isInteger(headerRow) || headerRow < 1)
  ) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_INVALID_HEADER_ROW,
      "The header row must be a positive whole number counted from 1.",
      { details: { headerRow } },
    );
  }
  const worksheet = requireWorksheet(workbook, sheetName);
  const origin: RegionOrigin =
    headerRow === undefined
      ? { kind: "detected-header" }
      : { kind: "declared-header" };

  if (column !== undefined && column.trim() !== "") {
    const header = findHeaderCell(worksheet, column, headerRow);
    if (!header) {
      throw columnNotFound(workbook, column, headerRow);
    }
    return [
      bindingForHeader(
        worksheet,
        header,
        column,
        await workbook.tables(),
        origin,
      ),
    ];
  }

  // Without a column to search for there is nothing to detect: the header is
  // the declared row, or the first row of the used range.
  const used = worksheet.usedRange;
  const effectiveHeaderRow = headerRow ?? used?.start.row ?? 1;
  return [
    new RangeBinding({
      body: {
        end: {
          column: used?.end.column ?? 0,
          row: used?.end.row ?? effectiveHeaderRow,
        },
        start: { column: used?.start.column ?? 0, row: effectiveHeaderRow + 1 },
      },
      headerRow: effectiveHeaderRow,
      origin,
      worksheet,
    }),
  ];
}

/**
 * Header search across sheets. `anchorText` locates the header row and
 * `columnText` is the column the caller will key on; for the split operation
 * they are the same string and this collapses to the old `findSplitHeader`
 * behavior exactly.
 */
function searchSheets(
  workbook: WorkbookModel,
  anchorText: string,
  columnText: string,
  tables: readonly WorkbookTableInfo[],
  stopAtFirst: boolean,
): DataRegion[] {
  const regions: DataRegion[] = [];
  for (const sheet of workbook.sheets) {
    const worksheet = workbook.worksheet(sheet.name);
    if (!worksheet) {
      continue;
    }
    const anchor = findHeaderCell(worksheet, anchorText, undefined);
    if (!anchor) {
      continue;
    }
    // The key column has to live on the header row the anchor found.
    const header =
      normalizeHeader(anchorText) === normalizeHeader(columnText)
        ? anchor
        : findHeaderCell(worksheet, columnText, anchor.row);
    if (!header) {
      continue;
    }
    regions.push(
      bindingForHeader(worksheet, header, columnText, tables, {
        kind: "detected-header",
      }),
    );
    if (stopAtFirst) {
      break;
    }
  }
  return regions;
}

export const resolveRegions: ResolveRegions = async (
  workbook: WorkbookModel,
  selector: RegionSelector,
  column?: string,
): Promise<readonly DataRegion[]> => {
  if (selector === "all-worksheets") {
    const columnText = requireColumn(column, "all-worksheets");
    const regions = searchSheets(
      workbook,
      columnText,
      columnText,
      await workbook.tables(),
      false,
    );
    if (regions.length === 0) {
      throw columnNotFound(workbook, columnText, undefined);
    }
    return regions;
  }
  if ("table" in selector) {
    return resolveTable(workbook, selector.table);
  }
  if ("range" in selector) {
    return resolveRange(workbook, selector.range);
  }
  if ("find" in selector) {
    const columnText = requireColumn(column, "find");
    const regions = searchSheets(
      workbook,
      selector.find,
      columnText,
      await workbook.tables(),
      true,
    );
    if (regions.length === 0) {
      throw columnNotFound(workbook, selector.find, undefined);
    }
    return regions;
  }
  return resolveSheet(workbook, selector.sheet, selector.headerRow, column);
};
