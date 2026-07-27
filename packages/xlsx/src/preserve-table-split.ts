import JSZip from "jszip";
import * as XLSX from "xlsx";

import type { ExcelTableDefinition } from "./excel-tables.js";

interface CellFragment {
  columnIndex: number;
  xml: string;
}

interface RowFragment {
  rowNumber: number;
  xml: string;
}

export interface PreserveExcelTableOptions {
  definition: ExcelTableDefinition;
  sourceRows: number[];
}

const CELL_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?c\b[^>]*?(?:\/\s*>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?c\s*>)/gu;
const ROW_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?row\b[^>]*?(?:\/\s*>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?row\s*>)/gu;
const SHEET_DATA_OPEN_PATTERN = /<(?:[A-Za-z_][\w.-]*:)?sheetData\b[^>]*>/u;

function xmlAttribute(xml: string, name: string): string | undefined {
  const match = new RegExp(`\\b${name}=(?:"([^"]*)"|'([^']*)')`, "u").exec(xml);
  return match?.[1] ?? match?.[2];
}

function qualifiedElementName(xml: string): string {
  const name = /^<([^\s/>]+)/u.exec(xml)?.[1];
  if (!name) {
    throw new Error("Could not determine an OOXML element name.");
  }
  return name;
}

function elementOpeningTag(xml: string): string {
  const end = xml.indexOf(">");
  if (end < 0) {
    throw new Error("Encountered an invalid OOXML element.");
  }
  return xml.slice(0, end + 1);
}

function parseCellFragments(rowXml: string): CellFragment[] {
  return [...rowXml.matchAll(CELL_PATTERN)].map((match) => {
    const xml = match[0];
    const reference = xmlAttribute(elementOpeningTag(xml), "r");
    if (!reference) {
      throw new Error("Encountered an OOXML cell without a reference.");
    }

    return {
      columnIndex: XLSX.utils.decode_cell(reference).c,
      xml,
    };
  });
}

function parseRowFragments(sheetDataXml: string): RowFragment[] {
  return [...sheetDataXml.matchAll(ROW_PATTERN)].map((match) => {
    const xml = match[0];
    const row = Number(xmlAttribute(elementOpeningTag(xml), "r"));
    if (!Number.isInteger(row) || row < 1) {
      throw new Error("Encountered an OOXML row without a valid row number.");
    }
    return { rowNumber: row, xml };
  });
}

function cellReference(xml: string): string {
  const reference = xmlAttribute(elementOpeningTag(xml), "r");
  if (!reference) {
    throw new Error("Encountered an OOXML cell without a reference.");
  }
  return reference;
}

function relocateCell(xml: string, destinationRow: number): string {
  const openingTag = elementOpeningTag(xml);
  const reference = cellReference(xml);
  const destinationReference = reference.replace(
    /\d+$/u,
    String(destinationRow),
  );
  const relocatedOpeningTag = openingTag.replace(
    /(\br=)(?:"[^"]*"|'[^']*')/u,
    `$1"${destinationReference}"`,
  );
  return `${relocatedOpeningTag}${xml.slice(openingTag.length)}`;
}

function clearCell(xml: string): string {
  const openingTag = elementOpeningTag(xml)
    .replace(/\s+t=(?:"[^"]*"|'[^']*')/u, "")
    .replace(/\s*\/?>$/u, " />");
  return openingTag;
}

function tableCells(
  rowXml: string | undefined,
  firstColumn: number,
  lastColumn: number,
): CellFragment[] {
  if (!rowXml) {
    return [];
  }
  return parseCellFragments(rowXml).filter(
    (cell) => cell.columnIndex >= firstColumn && cell.columnIndex <= lastColumn,
  );
}

function rewriteRow(
  rowXml: string | undefined,
  rowNumber: number,
  firstColumn: number,
  lastColumn: number,
  replacementCells: CellFragment[],
  rowElementName: string,
): string | undefined {
  const existingCells = rowXml ? parseCellFragments(rowXml) : [];
  const outsideCells = existingCells.filter(
    (cell) => cell.columnIndex < firstColumn || cell.columnIndex > lastColumn,
  );
  const cells = [...outsideCells, ...replacementCells].sort(
    (left, right) => left.columnIndex - right.columnIndex,
  );

  if (!rowXml && cells.length === 0) {
    return undefined;
  }

  if (!rowXml) {
    return `<${rowElementName} r="${rowNumber}">${cells
      .map((cell) => cell.xml)
      .join("")}</${rowElementName}>`;
  }

  const openingTag = elementOpeningTag(rowXml);
  const normalizedOpeningTag = openingTag.replace(/\s*\/>$/u, ">");
  const closingTag = `</${qualifiedElementName(rowXml)}>`;
  const body = rowXml.endsWith("/>")
    ? ""
    : rowXml.slice(openingTag.length, -closingTag.length);
  const nonCellBody = body.replace(CELL_PATTERN, "");

  return `${normalizedOpeningTag}${cells
    .map((cell) => cell.xml)
    .join("")}${nonCellBody}${closingTag}`;
}

function replaceElementReference(
  xml: string,
  localName: string,
  reference: string,
  required: boolean,
): string {
  const elementPattern = new RegExp(
    `<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b[^>]*>`,
    "u",
  );
  const match = elementPattern.exec(xml);
  if (!match) {
    if (required) {
      throw new Error(`The Excel Table is missing its ${localName} element.`);
    }
    return xml;
  }

  const openingTag = match[0];
  if (!/\bref=(?:"[^"]*"|'[^']*')/u.test(openingTag)) {
    if (required) {
      throw new Error(
        `The Excel Table ${localName} element has no range reference.`,
      );
    }
    return xml;
  }

  return `${xml.slice(0, match.index)}${openingTag.replace(
    /(\bref=)(?:"[^"]*"|'[^']*')/u,
    `$1"${reference}"`,
  )}${xml.slice(match.index + openingTag.length)}`;
}

function filterWorksheetXml(
  worksheetXml: string,
  definition: ExcelTableDefinition,
  sourceRows: number[],
): {
  tableDataReference: string;
  tableReference: string;
  worksheetXml: string;
} {
  const tableRange = XLSX.utils.decode_range(definition.range);
  const firstDataRow = tableRange.s.r + 1 + (definition.headerRow ? 1 : 0);
  const lastDataRow = tableRange.e.r + 1 - (definition.totalsRow ? 1 : 0);
  if (
    sourceRows.length === 0 ||
    sourceRows.some((row) => row < firstDataRow || row > lastDataRow)
  ) {
    throw new Error(
      `Excel Table "${definition.name}" received invalid source rows.`,
    );
  }

  const sheetDataOpeningMatch = SHEET_DATA_OPEN_PATTERN.exec(worksheetXml);
  if (!sheetDataOpeningMatch) {
    throw new Error(
      `Worksheet "${definition.sheet}" has no sheetData element.`,
    );
  }
  const sheetDataOpeningTag = sheetDataOpeningMatch[0];
  if (sheetDataOpeningTag.endsWith("/>")) {
    throw new Error(`Worksheet "${definition.sheet}" contains no rows.`);
  }
  const sheetDataElementName = qualifiedElementName(sheetDataOpeningTag);
  const sheetDataClosingTag = `</${sheetDataElementName}>`;
  const sheetDataStart =
    sheetDataOpeningMatch.index + sheetDataOpeningTag.length;
  const sheetDataEnd = worksheetXml.indexOf(
    sheetDataClosingTag,
    sheetDataStart,
  );
  if (sheetDataEnd < 0) {
    throw new Error(
      `Worksheet "${definition.sheet}" has invalid sheetData XML.`,
    );
  }

  const sheetDataXml = worksheetXml.slice(sheetDataStart, sheetDataEnd);
  const rowFragments = parseRowFragments(sheetDataXml);
  const rowByNumber = new Map(
    rowFragments.map((row) => [row.rowNumber, row.xml] as const),
  );
  const rowElementName = rowFragments[0]
    ? qualifiedElementName(rowFragments[0].xml)
    : sheetDataElementName.replace(/sheetData$/u, "row");
  const originalTableEndRow = tableRange.e.r + 1;
  const newLastDataRow = firstDataRow + sourceRows.length - 1;
  const newTableEndRow = newLastDataRow + (definition.totalsRow ? 1 : 0);
  const replacementCellsByRow = new Map<number, CellFragment[]>();

  for (let row = firstDataRow; row <= originalTableEndRow; row += 1) {
    replacementCellsByRow.set(
      row,
      tableCells(rowByNumber.get(row), tableRange.s.c, tableRange.e.c).map(
        (cell) => ({ ...cell, xml: clearCell(cell.xml) }),
      ),
    );
  }

  sourceRows.forEach((sourceRow, index) => {
    const destinationRow = firstDataRow + index;
    replacementCellsByRow.set(
      destinationRow,
      tableCells(
        rowByNumber.get(sourceRow),
        tableRange.s.c,
        tableRange.e.c,
      ).map((cell) => ({
        ...cell,
        xml: relocateCell(cell.xml, destinationRow),
      })),
    );
  });

  if (definition.totalsRow) {
    const originalTotalsRow = tableRange.e.r + 1;
    replacementCellsByRow.set(
      newTableEndRow,
      tableCells(
        rowByNumber.get(originalTotalsRow),
        tableRange.s.c,
        tableRange.e.c,
      ).map((cell) => ({
        ...cell,
        xml: relocateCell(cell.xml, newTableEndRow),
      })),
    );
  }

  for (const [rowNumber, replacementCells] of replacementCellsByRow) {
    const rewrittenRow = rewriteRow(
      rowByNumber.get(rowNumber),
      rowNumber,
      tableRange.s.c,
      tableRange.e.c,
      replacementCells,
      rowElementName,
    );
    if (rewrittenRow) {
      rowByNumber.set(rowNumber, rewrittenRow);
    } else {
      rowByNumber.delete(rowNumber);
    }
  }

  const rewrittenSheetData = [...rowByNumber]
    .sort(([left], [right]) => left - right)
    .map(([, rowXml]) => rowXml)
    .join("");
  const newWorksheetXml = `${worksheetXml.slice(
    0,
    sheetDataStart,
  )}${rewrittenSheetData}${worksheetXml.slice(sheetDataEnd)}`;
  const tableReference = XLSX.utils.encode_range({
    e: { c: tableRange.e.c, r: newTableEndRow - 1 },
    s: tableRange.s,
  });
  const tableDataReference = XLSX.utils.encode_range({
    e: { c: tableRange.e.c, r: newLastDataRow - 1 },
    s: tableRange.s,
  });

  return {
    tableDataReference,
    tableReference,
    worksheetXml: newWorksheetXml,
  };
}

export async function preserveWorkbookWithFilteredExcelTable(
  workbookBytes: Buffer,
  options: PreserveExcelTableOptions,
): Promise<Buffer> {
  const archive = await JSZip.loadAsync(workbookBytes);
  const worksheetEntry = archive.file(options.definition.worksheetPart);
  const tableEntry = archive.file(options.definition.tablePart);
  if (!worksheetEntry || !tableEntry) {
    throw new Error(
      `Excel Table "${options.definition.name}" is missing workbook package parts.`,
    );
  }

  const filtered = filterWorksheetXml(
    await worksheetEntry.async("text"),
    options.definition,
    options.sourceRows,
  );
  let tableXml = replaceElementReference(
    await tableEntry.async("text"),
    "table",
    filtered.tableReference,
    true,
  );
  tableXml = replaceElementReference(
    tableXml,
    "autoFilter",
    filtered.tableDataReference,
    false,
  );

  archive.file(options.definition.worksheetPart, filtered.worksheetXml);
  archive.file(options.definition.tablePart, tableXml);
  return archive.generateAsync({
    compression: "DEFLATE",
    type: "nodebuffer",
  });
}
