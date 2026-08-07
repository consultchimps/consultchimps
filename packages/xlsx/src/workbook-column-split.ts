import {
  mkdtemp,
  readFile,
  rename,
  rm,
  stat,
  writeFile,
} from "node:fs/promises";
import path from "node:path";

import {
  ConsultChimpsError,
  throwIfAborted,
  type OperationControlOptions,
  type OperationPlan,
  type OperationResult,
} from "@consultchimps/core";
import {
  ensureDirectory,
  ensureOutputAvailable,
  refuseInputOverwrite,
} from "@consultchimps/files";
import JSZip from "jszip";
import * as XLSX from "xlsx";

import {
  type ExcelTableDefinition,
  readExcelTableDefinitions,
  readWorkbookWorksheetParts,
} from "./excel-tables.js";
import { XLSX_ERRORS } from "./errors.js";
import { preserveWorkbookWithFilteredExcelTable } from "./preserve-table-split.js";
import { safeFilenameSegment, splitOutputPaths } from "./split-filenames.js";
import { pruneCalcChain } from "./tier1/calc-chain.js";
import { stripPivotParts } from "./tier1/pivot.js";
import { blankStaleCachedFormulas } from "./tier1/stale-values.js";
import { convertWorkbookToValuesWithReport } from "./values-only.js";

const SPLIT_OPERATION = "sheets.split-by-column";
const XLSX_MEDIA_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const XLSM_MEDIA_TYPE = "application/vnd.ms-excel.sheet.macroEnabled.12";
const ROW_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?row\b[^>]*?(?:\/\s*>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?row\s*>)/gu;
const CELL_REFERENCE_ATTRIBUTE_PATTERN = /(\br=)(["'])([A-Z]{1,3})\d+\2/gu;
const CELL_REFERENCE_CAPTURE_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?c\b[^>]*?\br=(?:"([^"]+)"|'([^']+)')/gu;
const SHEET_DATA_OPEN_PATTERN = /<(?:[A-Za-z_][\w.-]*:)?sheetData\b[^>]*>/u;
const FORMULA_XML_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?f\b([^>]*?)(?:\/\s*>|>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?f\s*>)/gu;
const NUMERIC_TEXT_PATTERN =
  /^[+-]?(?:(?:\d+(?:\.\d*)?)|(?:\.\d+))(?:e[+-]?\d+)?$/iu;

export type FullWorkbookSplitMetric =
  | "calcChainEntriesRemoved"
  | "formulaCellsBlankedForRemovedRows"
  | "formulaCellsConverted"
  | "formulaCellsWithoutCachedValues"
  | "groups"
  | "pivotTablesRemoved"
  | "inputFiles"
  | "inputRows"
  | "outputFiles"
  | "outputRows"
  | "rowsDeleted"
  | "sheetsCopiedUnchanged"
  | "sheetsFiltered"
  | "skippedRows"
  | "valuesOnly";

export interface FullWorkbookSplitOptions extends OperationControlOptions {
  column: string;
  filenamePrefix?: string | undefined;
  headerRow?: number | undefined;
  input: string;
  outputDirectory: string;
  overwrite?: boolean | undefined;
  strict?: boolean | undefined;
  values?: boolean | undefined;
}

export interface SplitSheetDetail {
  deletedRows: number;
  retainedRows: number;
  sheet: string;
}

export interface SplitOutputDetail {
  formulaCellsConverted: number;
  formulaCellsWithoutCachedValues: number;
  output: string;
  sheets: SplitSheetDetail[];
  value: string;
}

export interface FullWorkbookSplitSummary {
  column: string;
  copiedUnchangedSheets: string[];
  filteredSheets: string[];
  input: string;
  outputDirectory: string;
  valuesOnly: boolean;
}

export interface FullWorkbookSplitResult extends OperationResult<FullWorkbookSplitMetric> {
  outputs: SplitOutputDetail[];
  summary: FullWorkbookSplitSummary;
}

interface NormalizedValue {
  display: string;
  key: string;
}

type SplitGroup = NormalizedValue;

interface SheetAnalysis {
  firstDataRow: number;
  headerRow: number;
  lastDataRow: number;
  name: string;
  rowKeys: Map<number, string | undefined>;
  table: ExcelTableDefinition | undefined;
  worksheetPart: string;
}

interface ResolvedFullWorkbookSplit {
  absoluteInput: string;
  absoluteOutputDirectory: string;
  existingOutputs: Set<string>;
  extension: ".xlsm" | ".xlsx";
  groups: SplitGroup[];
  inputRows: number;
  mediaType: string;
  outputPaths: string[];
  sheets: SheetAnalysis[];
  skippedRows: number;
  unchangedSheets: string[];
  workbookBytes: Buffer;
}

function normalizeHeader(value: string): string {
  return value.normalize("NFKC").trim().toLowerCase();
}

function normalizeSplitValue(
  value: unknown,
  strict: boolean,
): NormalizedValue | undefined {
  if (value === null || value === undefined) {
    return undefined;
  }
  if (value instanceof Date) {
    const display = value.toISOString();
    return { display, key: `date:${display}` };
  }
  if (typeof value === "number") {
    if (!Number.isFinite(value)) {
      return undefined;
    }
    const display = String(Object.is(value, -0) ? 0 : value);
    return { display, key: `number:${display}` };
  }
  if (typeof value === "boolean") {
    return { display: String(value), key: `boolean:${String(value)}` };
  }

  const display = String(value).trim();
  if (!display) {
    return undefined;
  }
  if (strict) {
    return { display, key: `string:${String(value)}` };
  }
  if (NUMERIC_TEXT_PATTERN.test(display)) {
    const numericValue = Number(display);
    if (Number.isFinite(numericValue)) {
      return {
        display,
        key: `number:${String(Object.is(numericValue, -0) ? 0 : numericValue)}`,
      };
    }
  }
  return {
    display,
    key: `string:${display.normalize("NFKC").toLowerCase()}`,
  };
}

function worksheetCell(
  worksheet: XLSX.WorkSheet,
  row: number,
  column: number,
): XLSX.CellObject | undefined {
  return worksheet[XLSX.utils.encode_cell({ c: column, r: row })] as
    XLSX.CellObject | undefined;
}

function cellValue(cell: XLSX.CellObject | undefined): unknown {
  return cell?.v;
}

function findSplitHeader(
  worksheet: XLSX.WorkSheet,
  range: XLSX.Range,
  column: string,
  configuredHeaderRow: number | undefined,
): { columnIndex: number; rowIndex: number } | undefined {
  const target = normalizeHeader(column);
  const firstRow = configuredHeaderRow ? configuredHeaderRow - 1 : range.s.r;
  const lastRow = configuredHeaderRow ? firstRow : range.e.r;

  for (let rowIndex = firstRow; rowIndex <= lastRow; rowIndex += 1) {
    for (
      let columnIndex = range.s.c;
      columnIndex <= range.e.c;
      columnIndex += 1
    ) {
      const value = cellValue(worksheetCell(worksheet, rowIndex, columnIndex));
      if (value !== undefined && normalizeHeader(String(value)) === target) {
        return { columnIndex, rowIndex };
      }
    }
  }
  return undefined;
}

function findMatchingTable(
  definitions: ExcelTableDefinition[],
  sheet: string,
  headerRow: number,
  columnIndex: number,
  column: string,
): ExcelTableDefinition | undefined {
  const target = normalizeHeader(column);
  return definitions.find((definition) => {
    if (definition.sheet !== sheet || !definition.headerRow) {
      return false;
    }
    const range = XLSX.utils.decode_range(definition.range);
    const offset = columnIndex - range.s.c;
    return (
      range.s.r === headerRow &&
      offset >= 0 &&
      offset < definition.columns.length &&
      normalizeHeader(definition.columns[offset] ?? "") === target
    );
  });
}

function xmlAttribute(xml: string, name: string): string | undefined {
  const match = new RegExp(`\\b${name}=(?:"([^"]*)"|'([^']*)')`, "u").exec(xml);
  return match?.[1] ?? match?.[2];
}

interface WorksheetRowRemoval {
  /** References of the cells that left with their rows, e.g. `E7`. */
  deletedCells: Set<string>;
  /** Source row to destination row, for every row that survived. */
  renumberedRows: Map<number, number>;
  xml: string;
}

function removeWorksheetRows(
  worksheetXml: string,
  rowsToDelete: ReadonlySet<number>,
): WorksheetRowRemoval {
  // The deleted cells and the surviving rows' new numbers are the input the
  // Tier-1 calculation-chain pruner needs; collecting them here keeps the
  // single pass over the rows that already knows both.
  const deletedCells = new Set<string>();
  const renumberedRows = new Map<number, number>();
  if (rowsToDelete.size === 0) {
    return { deletedCells, renumberedRows, xml: worksheetXml };
  }
  const openMatch = SHEET_DATA_OPEN_PATTERN.exec(worksheetXml);
  if (!openMatch?.[0] || openMatch.index === undefined) {
    throw new Error("Worksheet package part does not contain sheetData.");
  }
  const openEnd = openMatch.index + openMatch[0].length;
  const qualifiedName = /^<([^\s/>]+)/u.exec(openMatch[0])?.[1];
  if (!qualifiedName) {
    throw new Error("Worksheet package part has invalid sheetData markup.");
  }
  const closeStart = worksheetXml.indexOf(`</${qualifiedName}>`, openEnd);
  if (closeStart < 0) {
    throw new Error("Worksheet package part has unclosed sheetData markup.");
  }
  const sheetData = worksheetXml.slice(openEnd, closeStart);
  let deletedBefore = 0;
  const filtered = sheetData.replace(ROW_PATTERN, (rowXml) => {
    const rowNumber = Number(
      xmlAttribute(rowXml.slice(0, rowXml.indexOf(">") + 1), "r"),
    );
    if (rowsToDelete.has(rowNumber)) {
      deletedBefore += 1;
      for (const cell of rowXml.matchAll(CELL_REFERENCE_CAPTURE_PATTERN)) {
        deletedCells.add((cell[1] ?? cell[2] ?? "").toUpperCase());
      }
      return "";
    }
    const destinationRow = rowNumber - deletedBefore;
    renumberedRows.set(rowNumber, destinationRow);
    if (destinationRow === rowNumber) {
      return rowXml;
    }
    return rowXml
      .replace(
        /^(<[\s\S]*?\brow\b[^>]*\br=)(["'])\d+\2/u,
        `$1"${destinationRow}"`,
      )
      .replace(CELL_REFERENCE_ATTRIBUTE_PATTERN, `$1$2$3${destinationRow}$2`);
  });
  return {
    deletedCells,
    renumberedRows,
    xml: `${worksheetXml.slice(0, openEnd)}${filtered}${worksheetXml.slice(closeStart)}`,
  };
}

/** The source rows a group's output drops from one analysed worksheet. */
function rowsRemovedFromSheet(
  sheet: SheetAnalysis,
  groupKey: string,
): Set<number> {
  const rowsToDelete = new Set<number>();
  for (let row = sheet.firstDataRow; row <= sheet.lastDataRow; row += 1) {
    if (sheet.rowKeys.get(row) !== groupKey) {
      rowsToDelete.add(row);
    }
  }
  return rowsToDelete;
}

interface PlainWorksheetFilter {
  bytes: Buffer;
  deletedCells: Map<string, Set<string>>;
  renumberedRows: Map<string, Map<number, number>>;
}

async function filterPlainWorksheets(
  workbookBytes: Buffer,
  sheets: SheetAnalysis[],
  groupKey: string,
): Promise<PlainWorksheetFilter> {
  const archive = await JSZip.loadAsync(workbookBytes);
  const removals = await Promise.all(
    sheets.map(async (sheet) => {
      const entry = archive.file(sheet.worksheetPart);
      if (!entry) {
        throw new Error(
          `Worksheet "${sheet.name}" is missing from the workbook package.`,
        );
      }
      const removal = removeWorksheetRows(
        await entry.async("text"),
        rowsRemovedFromSheet(sheet, groupKey),
      );
      archive.file(sheet.worksheetPart, removal.xml);
      return [sheet.worksheetPart, removal] as const;
    }),
  );
  return {
    bytes: await archive.generateAsync({
      compression: "DEFLATE",
      type: "nodebuffer",
    }),
    deletedCells: new Map(
      removals.map(([partName, removal]) => [partName, removal.deletedCells]),
    ),
    renumberedRows: new Map(
      removals.map(([partName, removal]) => [partName, removal.renumberedRows]),
    ),
  };
}

async function tableCanBeCompacted(
  workbookBytes: Buffer,
  sheet: SheetAnalysis,
): Promise<boolean> {
  const archive = await JSZip.loadAsync(workbookBytes);
  const entry = archive.file(sheet.worksheetPart);
  if (!entry) {
    return false;
  }
  const worksheetXml = await entry.async("text");
  for (const match of worksheetXml.matchAll(FORMULA_XML_PATTERN)) {
    const attributes = match[1] ?? "";
    const expression = (match[2] ?? "").replace(/"[^"]*"/gu, "");
    if (
      /\bt=(?:"(?:shared|array)"|'(?:shared|array)')/u.test(attributes) ||
      /(?<![A-Za-z0-9_.$])\$?[A-Za-z]{1,3}\$?\d+(?![\dA-Za-z_(])/u.test(
        expression,
      )
    ) {
      return false;
    }
  }
  return true;
}

function isMissingPathError(error: unknown): boolean {
  return (
    typeof error === "object" &&
    error !== null &&
    "code" in error &&
    error.code === "ENOENT"
  );
}

async function resolveFullWorkbookSplit(
  options: FullWorkbookSplitOptions,
): Promise<ResolvedFullWorkbookSplit> {
  const absoluteInput = path.resolve(options.input);
  const extension = path.extname(absoluteInput).toLowerCase();
  if (extension !== ".xlsx" && extension !== ".xlsm") {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_SPLIT_UNSUPPORTED_FILE,
      `Unsupported Excel workbook type "${extension || "(none)"}". Choose an .xlsx or .xlsm workbook.`,
      { details: { inputPath: absoluteInput } },
    );
  }
  if (options.headerRow !== undefined && options.headerRow < 1) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_INVALID_HEADER_ROW,
      "The header row must be a positive whole number counted from 1.",
      { details: { headerRow: options.headerRow } },
    );
  }

  let workbookBytes: Buffer;
  let workbook: XLSX.WorkBook;
  try {
    const inputStat = await stat(absoluteInput);
    if (!inputStat.isFile()) {
      throw new Error("The input path is not a file.");
    }
    workbookBytes = await readFile(absoluteInput);
    workbook = XLSX.read(workbookBytes, {
      bookVBA: extension === ".xlsm",
      cellDates: true,
      dense: false,
      type: "buffer",
    });
  } catch (error) {
    if (isMissingPathError(error)) {
      throw new ConsultChimpsError(
        XLSX_ERRORS.XLSX_INPUT_NOT_FOUND,
        `Input workbook was not found: ${absoluteInput}. Check the path and try again.`,
        { cause: error, details: { inputPath: absoluteInput } },
      );
    }
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_READ_FAILED,
      `Could not read workbook: ${absoluteInput}. Confirm that the file exists, is a valid unencrypted Excel workbook, and is not password protected.`,
      { cause: error, details: { inputPath: absoluteInput } },
    );
  }

  const [worksheetParts, tableDefinitions] = await Promise.all([
    readWorkbookWorksheetParts(workbookBytes),
    readExcelTableDefinitions(workbookBytes),
  ]).catch((error: unknown) => {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_READ_FAILED,
      `Could not inspect workbook structure: ${absoluteInput}. The file may be corrupted, encrypted, or contain invalid workbook XML.`,
      { cause: error, details: { inputPath: absoluteInput } },
    );
  });
  const worksheetPartByName = new Map(
    worksheetParts.map((sheet) => [sheet.name, sheet.worksheetPart] as const),
  );
  const groupsByKey = new Map<string, SplitGroup>();
  const sheets: SheetAnalysis[] = [];
  const unchangedSheets: string[] = [];
  let inputRows = 0;
  let skippedRows = 0;

  for (const sheetName of workbook.SheetNames) {
    const worksheet = workbook.Sheets[sheetName];
    const reference = worksheet?.["!ref"];
    const worksheetPart = worksheetPartByName.get(sheetName);
    if (!worksheet || !reference || !worksheetPart) {
      unchangedSheets.push(sheetName);
      continue;
    }
    const range = XLSX.utils.decode_range(reference);
    const header = findSplitHeader(
      worksheet,
      range,
      options.column,
      options.headerRow,
    );
    if (!header) {
      unchangedSheets.push(sheetName);
      continue;
    }
    const table = findMatchingTable(
      tableDefinitions,
      sheetName,
      header.rowIndex,
      header.columnIndex,
      options.column,
    );
    const tableRange = table ? XLSX.utils.decode_range(table.range) : undefined;
    const firstDataRow = header.rowIndex + 2;
    const lastDataRow = tableRange
      ? tableRange.e.r + 1 - (table?.totalsRow ? 1 : 0)
      : range.e.r + 1;

    const rowKeys = new Map<number, string | undefined>();
    if (lastDataRow >= firstDataRow) {
      for (let row = firstDataRow; row <= lastDataRow; row += 1) {
        inputRows += 1;
        const normalized = normalizeSplitValue(
          cellValue(worksheetCell(worksheet, row - 1, header.columnIndex)),
          options.strict === true,
        );
        rowKeys.set(row, normalized?.key);
        if (!normalized) {
          skippedRows += 1;
          continue;
        }
        if (!groupsByKey.has(normalized.key)) {
          groupsByKey.set(normalized.key, normalized);
        }
      }
    }
    sheets.push({
      firstDataRow,
      headerRow: header.rowIndex + 1,
      lastDataRow,
      name: sheetName,
      rowKeys,
      table,
      worksheetPart,
    });
  }

  if (sheets.length === 0) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_SPLIT_COLUMN_NOT_FOUND,
      `Column "${options.column}" was not found in any worksheet. Check the header text or provide the correct header row.`,
      {
        details: {
          availableWorksheets: workbook.SheetNames,
          column: options.column,
          headerRow: options.headerRow,
          inputPath: absoluteInput,
        },
      },
    );
  }
  const groups = [...groupsByKey.values()];
  if (groups.length === 0) {
    throw new ConsultChimpsError(
      XLSX_ERRORS.XLSX_SPLIT_NO_GROUPS,
      `Column "${options.column}" does not contain any non-blank values. Add at least one value and try again.`,
      { details: { column: options.column, inputPath: absoluteInput } },
    );
  }

  const absoluteOutputDirectory = path.resolve(options.outputDirectory);
  const prefix = options.filenamePrefix
    ? safeFilenameSegment(options.filenamePrefix, "split")
    : undefined;
  const outputPaths = splitOutputPaths(
    absoluteOutputDirectory,
    prefix,
    groups.map((group) => group.display),
    extension,
  );
  outputPaths.forEach((outputPath) =>
    refuseInputOverwrite(outputPath, [absoluteInput]),
  );
  const existingOutputs = new Set<string>();
  await Promise.all(
    outputPaths.map(async (outputPath) => {
      try {
        const outputStat = await stat(outputPath);
        if (!outputStat.isFile()) {
          throw new ConsultChimpsError(
            XLSX_ERRORS.XLSX_SPLIT_OUTPUT_NOT_FILE,
            `Output path exists but is not a file: ${outputPath}`,
            { details: { outputPath } },
          );
        }
        existingOutputs.add(outputPath);
      } catch (error) {
        if (!isMissingPathError(error)) {
          throw error;
        }
      }
    }),
  );

  return {
    absoluteInput,
    absoluteOutputDirectory,
    existingOutputs,
    extension,
    groups,
    inputRows,
    mediaType: extension === ".xlsm" ? XLSM_MEDIA_TYPE : XLSX_MEDIA_TYPE,
    outputPaths,
    sheets,
    skippedRows,
    unchangedSheets,
    workbookBytes,
  };
}

export async function planFullWorkbookSplit(
  options: FullWorkbookSplitOptions,
): Promise<OperationPlan<Exclude<FullWorkbookSplitMetric, "outputRows">>> {
  const resolved = await resolveFullWorkbookSplit(options);
  const warnings: string[] = [];
  if (resolved.existingOutputs.size > 0 && options.overwrite !== true) {
    warnings.push(
      `${resolved.existingOutputs.size} planned output file${resolved.existingOutputs.size === 1 ? " already exists" : "s already exist"}; executing without overwrite will fail.`,
    );
  }
  return {
    operation: SPLIT_OPERATION,
    inputs: [resolved.absoluteInput],
    outputs: resolved.outputPaths.map((outputPath) => ({
      exists: resolved.existingOutputs.has(outputPath),
      kind: "file",
      mediaType: resolved.mediaType,
      path: outputPath,
    })),
    warnings,
    metrics: {
      calcChainEntriesRemoved: 0,
      formulaCellsBlankedForRemovedRows: 0,
      formulaCellsConverted: 0,
      formulaCellsWithoutCachedValues: 0,
      groups: resolved.groups.length,
      inputFiles: 1,
      inputRows: resolved.inputRows,
      outputFiles: resolved.outputPaths.length,
      pivotTablesRemoved: 0,
      rowsDeleted: 0,
      sheetsCopiedUnchanged: resolved.unchangedSheets.length,
      sheetsFiltered: resolved.sheets.length,
      skippedRows: resolved.skippedRows,
      valuesOnly: options.values === true ? 1 : 0,
    },
  };
}

export async function splitFullWorkbookByColumn(
  options: FullWorkbookSplitOptions,
): Promise<FullWorkbookSplitResult> {
  throwIfAborted(options.signal, SPLIT_OPERATION);
  const resolved = await resolveFullWorkbookSplit(options);
  await Promise.all(
    resolved.outputPaths.map((outputPath) =>
      ensureOutputAvailable(outputPath, { overwrite: options.overwrite }),
    ),
  );
  await ensureDirectory(resolved.absoluteOutputDirectory);
  const transactionDirectory = await mkdtemp(
    path.join(resolved.absoluteOutputDirectory, ".consultchimps-split-"),
  );
  const stagedOutputs: string[] = [];
  const committedOutputs: string[] = [];
  const backups = new Map<string, string>();
  const outputDetails: SplitOutputDetail[] = [];
  const missingFormulaLocations = new Set<string>();
  const staleAggregateLocations = new Set<string>();
  const tableFallbackSheets = new Set<string>();
  let calcChainEntriesRemoved = 0;
  let formulaCellsBlankedForRemovedRows = 0;
  let formulaCellsConverted = 0;
  let formulaCellsWithoutCachedValues = 0;
  let outputRows = 0;
  let pivotTablesRemoved = 0;
  let rowsDeleted = 0;
  const worksheetNameByPart = new Map(
    (await readWorkbookWorksheetParts(resolved.workbookBytes)).map(
      (sheet) => [sheet.worksheetPart, sheet.name] as const,
    ),
  );

  try {
    for (const [index, group] of resolved.groups.entries()) {
      throwIfAborted(options.signal, SPLIT_OPERATION);
      let outputBytes: Uint8Array = resolved.workbookBytes;
      let outputFormulaCount = 0;
      let outputMissingFormulaCount = 0;
      const containsFilteredTable = resolved.sheets.some(
        (sheet) => sheet.table !== undefined,
      );
      // Tier-1 wiring: a values-only conversion bakes each formula's cached
      // result into the output, so any result computed over rows this group
      // does not receive is cleared first. This is the sole call site, placed
      // ahead of both conversion branches so it always sees source row
      // numbers; the conversion then reports the cleared cells as formulas
      // that lost their value.
      const deletedRowsByPart = new Map(
        resolved.sheets.map(
          (sheet) =>
            [
              sheet.worksheetPart,
              rowsRemovedFromSheet(sheet, group.key),
            ] as const,
        ),
      );
      if (options.values) {
        const staleValues = await blankStaleCachedFormulas(
          outputBytes,
          deletedRowsByPart,
        );
        outputBytes = staleValues.bytes;
        formulaCellsBlankedForRemovedRows += staleValues.blankedCells.length;
        staleValues.blankedCells.forEach((blanked) => {
          staleAggregateLocations.add(`${blanked.sheet}!${blanked.cell}`);
        });
      }
      if (options.values && containsFilteredTable) {
        const conversion = await convertWorkbookToValuesWithReport(outputBytes);
        outputBytes = conversion.bytes;
        outputFormulaCount = conversion.formulasConverted;
        outputMissingFormulaCount =
          conversion.formulasWithoutCachedValues.length;
        conversion.formulasWithoutCachedValues.forEach((missing) => {
          missingFormulaLocations.add(
            `${worksheetNameByPart.get(missing.worksheetPart) ?? missing.worksheetPart}!${missing.cell}`,
          );
        });
      }
      const compactTableSheets: SheetAnalysis[] = [];
      const plainSheets: SheetAnalysis[] = [];
      for (const sheet of resolved.sheets) {
        if (
          sheet.table &&
          (options.values ||
            (await tableCanBeCompacted(Buffer.from(outputBytes), sheet)))
        ) {
          compactTableSheets.push(sheet);
        } else {
          plainSheets.push(sheet);
          if (sheet.table) {
            tableFallbackSheets.add(sheet.name);
          }
        }
      }
      const sheets = resolved.sheets.map((sheet) => {
        const retainedRows = [...sheet.rowKeys.values()].filter(
          (key) => key === group.key,
        ).length;
        const deletedRows = sheet.rowKeys.size - retainedRows;
        outputRows += retainedRows;
        rowsDeleted += deletedRows;
        return { deletedRows, retainedRows, sheet: sheet.name };
      });

      for (const sheet of compactTableSheets) {
        outputBytes = await preserveWorkbookWithFilteredExcelTable(
          Buffer.from(outputBytes),
          {
            definition: sheet.table!,
            sourceRows: [...sheet.rowKeys]
              .filter(([, key]) => key === group.key)
              .map(([row]) => row),
            values: false,
            wholeRows: true,
          },
        );
      }
      const plainFilter = await filterPlainWorksheets(
        Buffer.from(outputBytes),
        plainSheets,
        group.key,
      );
      outputBytes = plainFilter.bytes;

      if (options.values && !containsFilteredTable) {
        const conversion = await convertWorkbookToValuesWithReport(outputBytes);
        outputBytes = conversion.bytes;
        outputFormulaCount = conversion.formulasConverted;
        outputMissingFormulaCount =
          conversion.formulasWithoutCachedValues.length;
        conversion.formulasWithoutCachedValues.forEach((missing) => {
          missingFormulaLocations.add(
            `${worksheetNameByPart.get(missing.worksheetPart) ?? missing.worksheetPart}!${missing.cell}`,
          );
        });
      }
      formulaCellsConverted += outputFormulaCount;
      formulaCellsWithoutCachedValues += outputMissingFormulaCount;

      // Tier-1 wiring: the calculation chain still names the cells the filter
      // deleted and the rows the survivors left. A values-only output has no
      // chain left to repair, so only the formula-preserving path prunes it.
      if (!options.values) {
        const prunedChain = await pruneCalcChain(
          outputBytes,
          plainFilter.deletedCells,
          plainFilter.renumberedRows,
        );
        outputBytes = prunedChain.bytes;
        calcChainEntriesRemoved += prunedChain.removedEntries;
      }

      // Tier-1 wiring: a pivot cache is a private copy of every source row, so
      // it would hand this group's recipient every other group's data. It
      // leaves with the rows it cached, on every output.
      const strippedPivots = await stripPivotParts(outputBytes);
      outputBytes = strippedPivots.bytes;
      pivotTablesRemoved += strippedPivots.removedPivotTables;

      const stagedOutput = path.join(
        transactionDirectory,
        `output-${String(index + 1).padStart(6, "0")}${resolved.extension}`,
      );
      await writeFile(stagedOutput, outputBytes);
      stagedOutputs.push(stagedOutput);
      outputDetails.push({
        formulaCellsConverted: outputFormulaCount,
        formulaCellsWithoutCachedValues: outputMissingFormulaCount,
        output: resolved.outputPaths[index]!,
        sheets,
        value: group.display,
      });
      options.onProgress?.({
        operation: SPLIT_OPERATION,
        stage: "staging-workbooks",
        completed: index + 1,
        total: resolved.groups.length,
        detail: `${path.basename(resolved.outputPaths[index]!)} (${sheets.map((sheet) => `${sheet.sheet}: kept ${sheet.retainedRows}, deleted ${sheet.deletedRows}`).join("; ")})`,
      });
    }

    for (const [index, outputPath] of resolved.outputPaths.entries()) {
      const stagedOutput = stagedOutputs[index];
      if (!stagedOutput) {
        continue;
      }
      if (resolved.existingOutputs.has(outputPath)) {
        const backupPath = path.join(
          transactionDirectory,
          `backup-${String(index + 1).padStart(6, "0")}${resolved.extension}`,
        );
        await rename(outputPath, backupPath);
        backups.set(outputPath, backupPath);
      }
      await rename(stagedOutput, outputPath);
      committedOutputs.push(outputPath);
      options.onProgress?.({
        operation: SPLIT_OPERATION,
        stage: "committing-outputs",
        completed: index + 1,
        total: resolved.outputPaths.length,
        detail: path.basename(outputPath),
      });
    }
  } catch (error) {
    const rollbackErrors: unknown[] = [];
    for (const outputPath of [...committedOutputs].reverse()) {
      try {
        await rm(outputPath, { force: true });
        const backupPath = backups.get(outputPath);
        if (backupPath) {
          await rename(backupPath, outputPath);
          backups.delete(outputPath);
        }
      } catch (rollbackError) {
        rollbackErrors.push(rollbackError);
      }
    }
    for (const [outputPath, backupPath] of backups) {
      try {
        await rename(backupPath, outputPath);
      } catch (rollbackError) {
        rollbackErrors.push(rollbackError);
      }
    }
    if (rollbackErrors.length > 0) {
      throw new ConsultChimpsError(
        XLSX_ERRORS.XLSX_SPLIT_ROLLBACK_FAILED,
        "The split failed and one or more output files could not be restored.",
        {
          cause: error,
          details: {
            outputPaths: resolved.outputPaths,
            rollbackErrors: rollbackErrors.map((rollbackError) =>
              rollbackError instanceof Error
                ? rollbackError.message
                : String(rollbackError),
            ),
          },
        },
      );
    }
    throw error;
  } finally {
    await rm(transactionDirectory, { force: true, recursive: true }).catch(
      () => undefined,
    );
  }

  const warnings: string[] = [];
  if (resolved.skippedRows > 0) {
    warnings.push(
      `Skipped ${resolved.skippedRows} row${resolved.skippedRows === 1 ? "" : "s"} with blank values in "${options.column}"; no blank-value workbook was created.`,
    );
  }
  if (formulaCellsWithoutCachedValues > 0) {
    const locations = [...missingFormulaLocations];
    const shown = locations.slice(0, 20).join(", ");
    warnings.push(
      `${formulaCellsWithoutCachedValues} formula cell${formulaCellsWithoutCachedValues === 1 ? " had" : "s had"} no cached value and became formatted blank cells in values-only output. Affected source locations: ${shown}${locations.length > 20 ? `, and ${locations.length - 20} more` : ""}. Open and recalculate the source workbook in Excel, save it, and rerun the split if these values are required.`,
    );
  }
  if (pivotTablesRemoved > 0) {
    warnings.push(
      `Removed ${pivotTablesRemoved} pivot table${pivotTablesRemoved === 1 ? "" : "s"}: their caches contained rows from other groups, and a cache travels inside the workbook whether or not the pivot is opened. Rebuild the pivot in Excel from each output's own rows if it is required.`,
    );
  }
  if (formulaCellsBlankedForRemovedRows > 0) {
    const locations = [...staleAggregateLocations];
    const shown = locations.slice(0, 20).join(", ");
    warnings.push(
      `${formulaCellsBlankedForRemovedRows} cached formula result${formulaCellsBlankedForRemovedRows === 1 ? " covered rows" : "s covered rows"} that are not part of this group and ${formulaCellsBlankedForRemovedRows === 1 ? "was" : "were"} cleared, so the values-only output shows a blank cell instead of a total computed over every group's rows. Affected locations: ${shown}${locations.length > 20 ? `, and ${locations.length - 20} more` : ""}. Recalculate them in Excel against the delivered rows if these values are required.`,
    );
  }
  if (tableFallbackSheets.size > 0) {
    warnings.push(
      `Excel Table${tableFallbackSheets.size === 1 ? "" : "s"} on ${[...tableFallbackSheets].join(", ")} contained formulas tied to row positions, so unmatched rows were removed without compacting the table range. The table's formatting and formulas were preserved; review the range in Excel before delivery.`,
    );
  }

  return {
    operation: SPLIT_OPERATION,
    artifacts: resolved.outputPaths.map((outputPath) => ({
      kind: "file",
      mediaType: resolved.mediaType,
      path: outputPath,
    })),
    warnings,
    metrics: {
      calcChainEntriesRemoved,
      formulaCellsBlankedForRemovedRows,
      formulaCellsConverted,
      formulaCellsWithoutCachedValues,
      groups: resolved.groups.length,
      inputFiles: 1,
      inputRows: resolved.inputRows,
      outputFiles: resolved.outputPaths.length,
      outputRows,
      pivotTablesRemoved,
      rowsDeleted,
      sheetsCopiedUnchanged: resolved.unchangedSheets.length,
      sheetsFiltered: resolved.sheets.length,
      skippedRows: resolved.skippedRows,
      valuesOnly: options.values === true ? 1 : 0,
    },
    outputs: outputDetails,
    summary: {
      column: options.column,
      copiedUnchangedSheets: resolved.unchangedSheets,
      filteredSheets: resolved.sheets.map((sheet) => sheet.name),
      input: resolved.absoluteInput,
      outputDirectory: resolved.absoluteOutputDirectory,
      valuesOnly: options.values === true,
    },
  };
}
