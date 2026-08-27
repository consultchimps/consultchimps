/**
 * Workbook inspection: describe a workbook's structure without producing
 * files. ADR 0002 makes this a first-class operation on every surface, with
 * `pptx.inspect-template` as the create-nothing precedent.
 *
 * The description answers "what is in this file, and what would an operation
 * see in it" — worksheets with their visibility and dimensions, the effective
 * header row and its columns, Excel Tables, named ranges, and a bounded sample
 * of each column's stored values. Samples are what mapping review and the tool
 * pickers need; unbounded values and inferred types are deliberately excluded,
 * so this module never reports a type it guessed and never carries a column's
 * whole contents.
 *
 * This module must stay free of node:fs and node:path imports so the byte
 * entry point can run in browsers.
 */
import {
  ConsultChimpsError,
  throwIfAborted,
  type AbortOutputContext,
  type OperationControlOptions,
  type OperationResult,
} from "@consultchimps/core";
import { type CellValue, uniqueHeaders } from "@consultchimps/tabular";
import * as XLSX from "xlsx";

import type { ExcelTableDefinition } from "./excel-tables.js";
import { XLSX_ERRORS } from "./errors.js";
import {
  BUILTIN_DEFINED_NAME_PREFIX,
  cellToPrimitive,
  findHeaderRow,
  getCell,
  INSPECT_OPERATION,
  lowercaseSet,
  parseNamedRangeRef,
  sheetVisibility,
  type ReadWorkbookOptions,
  type WorksheetVisibility,
} from "./shared.js";

/**
 * The hard ceiling on per-column sample values. ADR 0002 requires samples to
 * be bounded: a description is a summary a picker renders, never a copy of the
 * data. A caller may ask for fewer, never for more.
 */
export const MAX_COLUMN_SAMPLE_VALUES = 5;

export type DescribeWorkbookMetric =
  | "dataRows"
  | "excelTables"
  | "headerColumns"
  | "hiddenWorksheets"
  | "namedRanges"
  | "worksheets";

export interface DescribeWorkbookOptions
  extends ReadWorkbookOptions, OperationControlOptions {
  /**
   * Distinct sample values to collect per column, from 0 to
   * `MAX_COLUMN_SAMPLE_VALUES`. Defaults to the maximum.
   */
  sampleValues?: number | undefined;
}

/** One column of a worksheet's effective header row. */
export interface WorkbookColumnDescription {
  /** The header as the union readers would spell it, blanks filled in. */
  header: string;
  /** Zero-based position within the header row. */
  index: number;
  /**
   * The first few distinct non-empty stored values below the header, in the
   * order the rows carry them. Bounded by the sample limit; never the whole
   * column, and never a type this package inferred.
   */
  sampleValues: CellValue[];
}

export interface WorkbookSheetDescription {
  name: string;
  visibility: WorksheetVisibility;
  /** Rows in the worksheet's used range, header row included; 0 when empty. */
  rowCount: number;
  /** Columns in the worksheet's used range; 0 when empty. */
  columnCount: number;
  /**
   * The one-based effective header row, or undefined when the worksheet holds
   * no values for one to be found in.
   */
  headerRow: number | undefined;
  /** The header preview: one entry per column of the effective header row. */
  columns: WorkbookColumnDescription[];
  /** Non-empty rows below the header row. */
  dataRowCount: number;
}

export interface WorkbookExcelTableDescription {
  name: string;
  range: string;
  sheet: string;
  /** The column names the table part declares, in table order. */
  headers: string[];
}

export interface WorkbookNamedRangeDescription {
  name: string;
  /** The range this name points at, as the workbook stores it. */
  ref: string;
  sheet: string;
}

export interface WorkbookDescription {
  /** The workbook this description was read from: a filename or input name. */
  source: string;
  sheets: WorkbookSheetDescription[];
  excelTables: WorkbookExcelTableDescription[];
  namedRanges: WorkbookNamedRangeDescription[];
}

/**
 * The outcome of a workbook inspection: the structured operation result every
 * completed operation reports, plus the description it summarizes. The two
 * travel side by side for the same reason `ByteOperationOutcome` keeps
 * `outputs` beside `result` — metrics are counts, and sheet names, headers and
 * sample values are not.
 */
export interface WorkbookDescriptionOutcome {
  description: WorkbookDescription;
  result: OperationResult<DescribeWorkbookMetric>;
}

/**
 * Validate the sample bound. Silently clamping a caller's 50 to 5 would make
 * the same call mean different things on either side of the cap, so an
 * out-of-range request is a stable refusal instead.
 */
function resolveSampleLimit(requested: number | undefined): number {
  if (requested === undefined) {
    return MAX_COLUMN_SAMPLE_VALUES;
  }
  if (
    !Number.isInteger(requested) ||
    requested < 0 ||
    requested > MAX_COLUMN_SAMPLE_VALUES
  ) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_INVALID_SAMPLE_LIMIT,
      `The sample value count must be a whole number from 0 to ${MAX_COLUMN_SAMPLE_VALUES}.`,
      {
        details: {
          maximum: MAX_COLUMN_SAMPLE_VALUES,
          sampleValues: requested,
        },
      },
    );
  }
  return requested;
}

/**
 * A sample value's identity for the distinctness test. The type is part of the
 * key so the number 1 and the text "1" — which Excel very much distinguishes,
 * and which a mapping review needs to see separately — do not collapse into
 * one sample.
 */
function sampleKey(value: CellValue): string {
  return `${typeof value}:${String(value)}`;
}

function isEmptyValue(value: CellValue): boolean {
  return value === null || value === "";
}

/**
 * Describe one worksheet: its dimensions, its effective header row, and a
 * bounded sample of each column's values.
 *
 * The scan walks rows in worksheet order and stops as soon as every column has
 * filled its sample quota, so a wide workbook costs a few rows rather than a
 * full read. Because the order is the row order and the bound is fixed, the
 * same worksheet always yields the same samples.
 */
function describeWorksheet(
  workbook: XLSX.WorkBook,
  sheetName: string,
  configuredHeaderRow: number | undefined,
  sampleLimit: number,
): WorkbookSheetDescription {
  const visibility = sheetVisibility(workbook, sheetName);
  const worksheet = workbook.Sheets[sheetName];
  const reference = worksheet?.["!ref"];

  if (!worksheet || !reference) {
    return {
      name: sheetName,
      visibility,
      rowCount: 0,
      columnCount: 0,
      headerRow: undefined,
      columns: [],
      dataRowCount: 0,
    };
  }

  const range = XLSX.utils.decode_range(reference);
  const rowCount = range.e.r - range.s.r + 1;
  const columnCount = range.e.c - range.s.c + 1;
  const headerRowIndex = findHeaderRow(worksheet, range, configuredHeaderRow);

  if (
    headerRowIndex === undefined ||
    headerRowIndex < range.s.r ||
    headerRowIndex > range.e.r
  ) {
    return {
      name: sheetName,
      visibility,
      rowCount,
      columnCount,
      headerRow: undefined,
      columns: [],
      dataRowCount: 0,
    };
  }

  const rawHeaders: Array<string | null> = [];
  for (
    let columnIndex = range.s.c;
    columnIndex <= range.e.c;
    columnIndex += 1
  ) {
    const value = cellToPrimitive(
      getCell(worksheet, headerRowIndex, columnIndex),
    );
    rawHeaders.push(value === null ? null : String(value));
  }
  // The same de-duplication the readers apply, so a header preview shows the
  // names an operation would actually produce rather than the raw cells.
  const headers = uniqueHeaders(rawHeaders);

  const samples = headers.map(() => [] as CellValue[]);
  const seen = headers.map(() => new Set<string>());
  let dataRowCount = 0;
  let satisfiedColumns = sampleLimit === 0 ? headers.length : 0;

  for (
    let rowIndex = headerRowIndex + 1;
    rowIndex <= range.e.r;
    rowIndex += 1
  ) {
    const values = headers.map((_, index) =>
      cellToPrimitive(getCell(worksheet, rowIndex, range.s.c + index)),
    );
    if (values.every(isEmptyValue)) {
      continue;
    }
    dataRowCount += 1;

    if (satisfiedColumns >= headers.length) {
      // Every column's samples are complete; keep counting rows, which is a
      // cell-free walk, but stop reading values.
      continue;
    }

    values.forEach((value, index) => {
      const columnSamples = samples[index]!;
      if (isEmptyValue(value) || columnSamples.length >= sampleLimit) {
        return;
      }
      const key = sampleKey(value);
      const columnSeen = seen[index]!;
      if (columnSeen.has(key)) {
        return;
      }
      columnSeen.add(key);
      columnSamples.push(value);
      if (columnSamples.length === sampleLimit) {
        satisfiedColumns += 1;
      }
    });
  }

  return {
    name: sheetName,
    visibility,
    rowCount,
    columnCount,
    headerRow: headerRowIndex + 1,
    columns: headers.map((header, index) => ({
      header,
      index,
      sampleValues: samples[index]!,
    })),
    dataRowCount,
  };
}

function describeExcelTables(
  definitions: readonly ExcelTableDefinition[],
  describedSheets: ReadonlySet<string>,
): WorkbookExcelTableDescription[] {
  return definitions
    .filter((definition) => describedSheets.has(definition.sheet))
    .map((definition) => ({
      name: definition.name,
      range: definition.range,
      sheet: definition.sheet,
      // The declared column names, not the cells: a table with no data rows
      // still has headers, and a picker needs to offer them.
      headers: [...definition.columns],
    }));
}

function describeNamedRanges(
  workbook: XLSX.WorkBook,
  describedSheets: ReadonlySet<string>,
): WorkbookNamedRangeDescription[] {
  const ranges: WorkbookNamedRangeDescription[] = [];

  for (const definedName of workbook.Workbook?.Names ?? []) {
    if (
      !definedName.Name ||
      definedName.Name.startsWith(BUILTIN_DEFINED_NAME_PREFIX)
    ) {
      continue;
    }
    const parsed = parseNamedRangeRef(definedName.Ref ?? "");
    if (!parsed || !describedSheets.has(parsed.sheet)) {
      continue;
    }
    ranges.push({
      name: definedName.Name,
      ref: parsed.range,
      sheet: parsed.sheet,
    });
  }

  return ranges;
}

/**
 * The worksheets this description covers, in workbook order.
 *
 * `includeHiddenSheets` and `sheets` mean exactly what they mean to every
 * other reader in this package — an option that reads the same everywhere is
 * worth more than an inspection-specific default. Naming a worksheet the
 * workbook does not have is a refusal rather than an empty answer, because a
 * picker asking about a sheet that is not there has a mistake to report.
 */
function selectSheets(
  workbook: XLSX.WorkBook,
  options: DescribeWorkbookOptions,
): { hiddenExcluded: number; names: string[] } {
  if (workbook.SheetNames.length === 0) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_NO_SHEETS,
      "The workbook does not contain any worksheets.",
      { details: { availableWorksheets: [] } },
    );
  }

  const selected = lowercaseSet(options.sheets);
  if (selected) {
    const available = new Set(
      workbook.SheetNames.map((name) => name.toLocaleLowerCase()),
    );
    const missing = (options.sheets ?? []).filter(
      (name) => !available.has(name.toLocaleLowerCase()),
    );
    if (missing.length > 0) {
      throw new ConsultChimpsError(
        XLSX_ERRORS.XLSX_WORKSHEET_NOT_FOUND,
        `Worksheet "${missing[0]}" was not found in the workbook.`,
        {
          details: {
            availableWorksheets: workbook.SheetNames,
            missingWorksheets: missing,
          },
        },
      );
    }
  }

  const names: string[] = [];
  let hiddenExcluded = 0;

  for (const sheetName of workbook.SheetNames) {
    if (selected && !selected.has(sheetName.toLocaleLowerCase())) {
      continue;
    }
    if (
      options.includeHiddenSheets !== true &&
      sheetVisibility(workbook, sheetName) !== "visible"
    ) {
      hiddenExcluded += 1;
      continue;
    }
    names.push(sheetName);
  }

  return { hiddenExcluded, names };
}

/**
 * The conditions an inspection reports as warnings. Order is fixed rather than
 * derived from the workbook, so identical inputs produce an identical result.
 */
function descriptionWarnings(
  description: WorkbookDescription,
  hiddenExcluded: number,
): string[] {
  const warnings: string[] = [];

  if (hiddenExcluded > 0) {
    warnings.push(
      `${hiddenExcluded} worksheet${
        hiddenExcluded === 1 ? " is" : "s are"
      } hidden and ${
        hiddenExcluded === 1 ? "was" : "were"
      } not described. Include hidden worksheets to describe ${
        hiddenExcluded === 1 ? "it" : "them"
      }.`,
    );
  }

  const withoutHeaderRow = description.sheets
    .filter((sheet) => sheet.headerRow === undefined)
    .map((sheet) => sheet.name);
  if (withoutHeaderRow.length > 0) {
    warnings.push(
      `No header row was found in ${withoutHeaderRow
        .map((name) => `"${name}"`)
        .join(
          ", ",
        )}. An operation that matches columns by header would find nothing to match in ${
        withoutHeaderRow.length === 1 ? "it" : "them"
      }.`,
    );
  }

  if (description.sheets.length === 0) {
    warnings.push(
      "No worksheets matched the selection, so the description is empty.",
    );
  }

  return warnings;
}

/** Present a description as the structured result every operation reports. */
export function workbookDescriptionResult(
  description: WorkbookDescription,
  hiddenExcluded: number,
): OperationResult<DescribeWorkbookMetric> {
  return {
    operation: INSPECT_OPERATION,
    // An inspection creates nothing, so it has no artifacts. The structure
    // itself travels beside this result: metrics are counts, and names are not.
    artifacts: [],
    warnings: descriptionWarnings(description, hiddenExcluded),
    metrics: {
      dataRows: description.sheets.reduce(
        (total, sheet) => total + sheet.dataRowCount,
        0,
      ),
      excelTables: description.excelTables.length,
      headerColumns: description.sheets.reduce(
        (total, sheet) => total + sheet.columns.length,
        0,
      ),
      hiddenWorksheets: description.sheets.filter(
        (sheet) => sheet.visibility !== "visible",
      ).length,
      namedRanges: description.namedRanges.length,
      worksheets: description.sheets.length,
    },
  };
}

/**
 * Describe a parsed workbook. Both surfaces call this with the same parsed
 * workbook and the same table definitions, which is what makes the file and
 * byte descriptions of one workbook structurally identical.
 *
 * The work is a bounded read of already-parsed cells, so the abort checks sit
 * at the worksheet boundaries rather than inside the scan: cheap as this is,
 * every operation honours `signal` and `onProgress` so a caller never has to
 * ask which ones do.
 */
export function describeParsedWorkbook(
  workbook: XLSX.WorkBook,
  definitions: readonly ExcelTableDefinition[],
  source: string,
  options: DescribeWorkbookOptions = {},
  outputContext: AbortOutputContext = "files",
): WorkbookDescriptionOutcome {
  throwIfAborted(options.signal, INSPECT_OPERATION, outputContext);
  const sampleLimit = resolveSampleLimit(options.sampleValues);
  const { hiddenExcluded, names } = selectSheets(workbook, options);

  const sheets: WorkbookSheetDescription[] = [];
  for (const [index, sheetName] of names.entries()) {
    throwIfAborted(options.signal, INSPECT_OPERATION, outputContext);
    sheets.push(
      describeWorksheet(workbook, sheetName, options.headerRow, sampleLimit),
    );
    options.onProgress?.({
      operation: INSPECT_OPERATION,
      stage: "describing-worksheets",
      completed: index + 1,
      total: names.length,
      detail: sheetName,
    });
  }

  const describedSheets = new Set(names);
  const description: WorkbookDescription = {
    source,
    sheets,
    excelTables: describeExcelTables(definitions, describedSheets),
    namedRanges: describeNamedRanges(workbook, describedSheets),
  };

  throwIfAborted(options.signal, INSPECT_OPERATION, outputContext);
  options.onProgress?.({
    operation: INSPECT_OPERATION,
    stage: "describing-structures",
    completed: 1,
    total: 1,
  });

  return {
    description,
    result: workbookDescriptionResult(description, hiddenExcluded),
  };
}
