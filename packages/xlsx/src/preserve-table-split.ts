/**
 * The preserved Excel Table split: keep the whole source package and replace
 * only the selected table's rows.
 *
 * This is the one split mode still expressed as an XML rewrite rather than
 * through the layered engine. It is kept deliberately: its contract is
 * *refusal*, not repair. When a row that carries a hand-written A1 formula
 * would move, it raises `XLSX_SPLIT_PRESERVE_FORMULA` and produces nothing,
 * so the caller decides between converting to values and splitting without
 * preserving the workbook. Re-expressing it on `TableBinding` would silently
 * relocate those formulas instead, which is a different promise; making that
 * change belongs with the phase that decides it, not with this migration.
 *
 * Phase 1 did retire the module's dead half: a cell-only rewrite mode that no
 * caller selected, which cleared table cells in place rather than removing
 * whole rows.
 */
import { ConsultChimpsError } from "@consultchimps/core";
import * as XLSX from "xlsx";

import type { ExcelTableDefinition } from "./excel-tables.js";
import { XLSX_ERRORS } from "./errors.js";
import { WorkbookPackage } from "./package/index.js";

interface RowFragment {
  rowNumber: number;
  xml: string;
}

export interface PreserveExcelTableOptions {
  definition: ExcelTableDefinition;
  sourceRows: number[];
  values?: boolean | undefined;
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

const FORMULA_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?f\b[^>]*(?:\/\s*>|>(?<expression>[\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?f\s*>)/u;
// An A1-style cell reference outside a quoted string. The negative lookahead
// keeps function names such as LOG10( from matching.
const A1_REFERENCE_PATTERN =
  /(?<![A-Za-z0-9_.$])\$?[A-Za-z]{1,3}\$?\d+(?![\dA-Za-z_(])/u;

function assertRelocatableFormula(xml: string, destinationRow: number): void {
  const match = FORMULA_PATTERN.exec(xml);
  if (!match) {
    return;
  }

  const formulaOpeningTag = elementOpeningTag(match[0]);
  const formulaType = xmlAttribute(formulaOpeningTag, "t");
  const expression = (match.groups?.expression ?? "").replace(
    /"[^"]*"/gu,
    '""',
  );
  const positionDependent =
    formulaType === "shared" ||
    formulaType === "array" ||
    A1_REFERENCE_PATTERN.test(expression);
  if (!positionDependent) {
    return;
  }

  const reference = cellReference(xml);
  throw new ConsultChimpsError(
    XLSX_ERRORS.XLSX_SPLIT_PRESERVE_FORMULA,
    `Cell ${reference} contains a formula with cell references, and its row would move during a preserved split, which would leave the formula pointing at the wrong rows. Convert the formula to structured table references, or run the split again without preserving the workbook.`,
    {
      details: { cell: reference, destinationRow },
    },
  );
}

function relocateCell(
  xml: string,
  destinationRow: number,
  values: boolean,
): string {
  const preparedXml = values ? xml.replace(FORMULA_PATTERN, "") : xml;
  const openingTag = elementOpeningTag(preparedXml);
  const reference = cellReference(preparedXml);
  const sourceRow = Number(/\d+$/u.exec(reference)?.[0]);
  if (sourceRow === destinationRow) {
    return preparedXml;
  }

  assertRelocatableFormula(preparedXml, destinationRow);
  const destinationReference = reference.replace(
    /\d+$/u,
    String(destinationRow),
  );
  const relocatedOpeningTag = openingTag.replace(
    /(\br=)(?:"[^"]*"|'[^']*')/u,
    `$1"${destinationReference}"`,
  );
  return `${relocatedOpeningTag}${preparedXml.slice(openingTag.length)}`;
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

function filterWholeWorksheetRows(
  worksheetXml: string,
  definition: ExcelTableDefinition,
  sourceRows: number[],
  values: boolean,
): {
  tableDataReference: string;
  tableReference: string;
  worksheetXml: string;
} {
  const tableRange = XLSX.utils.decode_range(definition.range);
  const firstDataRow = tableRange.s.r + 2;
  const originalLastDataRow =
    tableRange.e.r + 1 - (definition.totalsRow ? 1 : 0);
  const originalTableEndRow = tableRange.e.r + 1;
  if (
    sourceRows.some((row) => row < firstDataRow || row > originalLastDataRow)
  ) {
    throw new Error(
      `Excel Table "${definition.name}" received invalid source rows.`,
    );
  }

  const openingMatch = SHEET_DATA_OPEN_PATTERN.exec(worksheetXml);
  if (!openingMatch) {
    throw new Error(
      `Worksheet "${definition.sheet}" has no sheetData element.`,
    );
  }
  const sheetDataOpeningTag = openingMatch[0];
  const sheetDataElementName = qualifiedElementName(sheetDataOpeningTag);
  const sheetDataClosingTag = `</${sheetDataElementName}>`;
  const sheetDataStart = openingMatch.index + sheetDataOpeningTag.length;
  const sheetDataEnd = worksheetXml.indexOf(
    sheetDataClosingTag,
    sheetDataStart,
  );
  if (sheetDataEnd < 0) {
    throw new Error(
      `Worksheet "${definition.sheet}" has invalid sheetData XML.`,
    );
  }

  const rowFragments = parseRowFragments(
    worksheetXml.slice(sheetDataStart, sheetDataEnd),
  );
  const rowByNumber = new Map(
    rowFragments.map((row) => [row.rowNumber, row.xml] as const),
  );
  const rowElementName = rowFragments[0]
    ? qualifiedElementName(rowFragments[0].xml)
    : sheetDataElementName.replace(/sheetData$/u, "row");
  const sourceXmlByRow = new Map(
    sourceRows.map((row) => [row, rowByNumber.get(row)] as const),
  );
  const newLastDataRow = sourceRows.length
    ? firstDataRow + sourceRows.length - 1
    : firstDataRow - 1;
  const newTableEndRow = definition.totalsRow
    ? firstDataRow + sourceRows.length
    : Math.max(newLastDataRow, tableRange.s.r + 1);

  for (let row = firstDataRow; row <= originalTableEndRow; row += 1) {
    rowByNumber.delete(row);
  }
  sourceRows.forEach((sourceRow, index) => {
    const destinationRow = firstDataRow + index;
    const sourceXml = sourceXmlByRow.get(sourceRow);
    if (!sourceXml) {
      return;
    }
    const relocated = sourceXml
      .replace(
        /^(<[^\s/>]*row\b[^>]*\br=)(?:"[^"]*"|'[^']*')/u,
        `$1"${destinationRow}"`,
      )
      .replace(CELL_PATTERN, (cellXml) =>
        relocateCell(cellXml, destinationRow, values),
      );
    rowByNumber.set(
      destinationRow,
      relocated.replace(/<row\b/u, `<${rowElementName}`),
    );
  });

  if (definition.totalsRow) {
    const totalsRow = rowFragments.find(
      (row) => row.rowNumber === originalTableEndRow,
    )?.xml;
    if (totalsRow) {
      const relocatedTotals = totalsRow
        .replace(
          /^(<[^\s/>]*row\b[^>]*\br=)(?:"[^"]*"|'[^']*')/u,
          `$1"${newTableEndRow}"`,
        )
        .replace(CELL_PATTERN, (cellXml) =>
          relocateCell(cellXml, newTableEndRow, values),
        );
      rowByNumber.set(
        newTableEndRow,
        relocatedTotals.replace(/<row\b/u, `<${rowElementName}`),
      );
    }
  }

  const rewrittenSheetData = [...rowByNumber]
    .sort(([left], [right]) => left - right)
    .map(([, rowXml]) => rowXml)
    .join("");
  const tableReference = XLSX.utils.encode_range({
    e: { c: tableRange.e.c, r: newTableEndRow - 1 },
    s: tableRange.s,
  });
  const tableDataReference = XLSX.utils.encode_range({
    e: { c: tableRange.e.c, r: Math.max(newLastDataRow - 1, tableRange.s.r) },
    s: tableRange.s,
  });
  return {
    tableDataReference,
    tableReference,
    worksheetXml: `${worksheetXml.slice(0, sheetDataStart)}${rewrittenSheetData}${worksheetXml.slice(sheetDataEnd)}`,
  };
}

export async function preserveWorkbookWithFilteredExcelTable(
  workbookBytes: Uint8Array,
  options: PreserveExcelTableOptions,
): Promise<Uint8Array> {
  const workbookPackage = await WorkbookPackage.load(workbookBytes);
  const worksheetXml = workbookPackage.readText(
    options.definition.worksheetPart,
  );
  const tableXmlSource = workbookPackage.readText(options.definition.tablePart);
  if (worksheetXml === undefined || tableXmlSource === undefined) {
    throw new Error(
      `Excel Table "${options.definition.name}" is missing workbook package parts.`,
    );
  }

  const filtered = filterWholeWorksheetRows(
    worksheetXml,
    options.definition,
    options.sourceRows,
    options.values === true,
  );
  let tableXml = replaceElementReference(
    tableXmlSource,
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

  workbookPackage.writeText(
    options.definition.worksheetPart,
    filtered.worksheetXml,
  );
  workbookPackage.writeText(options.definition.tablePart, tableXml);
  return workbookPackage.save();
}
