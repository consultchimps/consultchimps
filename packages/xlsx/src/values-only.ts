import { hasCachedValueElement } from "./model/xml.js";
import {
  forEachWorkbookSheet,
  tagAttribute,
  WorkbookPackage,
} from "./package/index.js";

const CELL_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?c\b[^>]*?(?:\/\s*>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?c\s*>)/gu;
const CELL_FORMULA_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?f\b[^>]*(?:\/\s*>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?f\s*>)/gu;
const TABLE_FORMULA_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?(?:calculatedColumnFormula|totalsRowFormula)\b[^>]*(?:\/\s*>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?(?:calculatedColumnFormula|totalsRowFormula)\s*>)/gu;
const CELL_REFERENCE_PATTERN = /\br=(?:"([^"]+)"|'([^']+)')/u;
const WORKSHEET_PART_PATTERN = /^xl\/worksheets\/[^/]+\.xml$/iu;
const TABLE_PART_PATTERN = /^xl\/tables\/[^/]+\.xml$/iu;
const CALC_CHAIN_PART = "xl/calcChain.xml";
const WORKBOOK_PART = "xl/workbook.xml";

export interface MissingCachedFormula {
  cell: string;
  worksheetPart: string;
  /** `Sheet!B4`, or `part!B4` for a part the workbook does not list. */
  location: string;
}

export interface ValuesOnlyConversion {
  bytes: Uint8Array;
  formulasConverted: number;
  formulasWithoutCachedValues: MissingCachedFormula[];
}

export function removeWorksheetFormulas(
  worksheetXml: string,
  worksheetPart: string,
): {
  formulasConverted: number;
  formulasWithoutCachedValues: Omit<MissingCachedFormula, "location">[];
  xml: string;
} {
  let formulasConverted = 0;
  const formulasWithoutCachedValues: Omit<MissingCachedFormula, "location">[] =
    [];
  const xml = worksheetXml.replace(CELL_PATTERN, (cellXml) => {
    if (!CELL_FORMULA_PATTERN.test(cellXml)) {
      return cellXml;
    }

    CELL_FORMULA_PATTERN.lastIndex = 0;
    formulasConverted += 1;
    if (!hasCachedValueElement(cellXml)) {
      const reference = CELL_REFERENCE_PATTERN.exec(cellXml);
      formulasWithoutCachedValues.push({
        cell: reference?.[1] ?? reference?.[2] ?? "unknown cell",
        worksheetPart,
      });
    }
    return cellXml.replace(CELL_FORMULA_PATTERN, "");
  });

  return { formulasConverted, formulasWithoutCachedValues, xml };
}

/**
 * Replace worksheet formulas with their cached values without rebuilding any
 * cells. Editing the OOXML parts directly preserves cell styles, number
 * formats, row heights, column widths, tables, and the rest of the workbook
 * package byte-for-byte apart from formula and calculation-chain metadata.
 */
/** Each worksheet part's sheet name, as the workbook part lists them. */
export function worksheetNamesByPart(
  workbookPackage: WorkbookPackage,
): Map<string, string> {
  const targets = new Map(
    workbookPackage
      .relationshipsOf(WORKBOOK_PART)
      .map(
        (relationship) =>
          [
            relationship.id,
            workbookPackage.resolvePart(WORKBOOK_PART, relationship.target),
          ] as const,
      ),
  );
  const names = new Map<string, string>();
  const xml = workbookPackage.readText(WORKBOOK_PART);
  if (xml === undefined) return names;
  forEachWorkbookSheet(xml, WORKBOOK_PART, (tag) => {
    const name = tagAttribute(tag, "name");
    const part = targets.get(tagAttribute(tag, "id") ?? "");
    if (name !== undefined && part !== undefined) names.set(part, name);
  });
  return names;
}

/** Worksheet parts the values conversion rewrites. */
export function isConvertedWorksheetPart(partName: string): boolean {
  return WORKSHEET_PART_PATTERN.test(partName);
}

/**
 * The rest of a values conversion once the worksheets are done: table column
 * formulas removed, and the calculation chain, which lists formulas, dropped,
 * since a workbook with no formulas left has nothing to calculate.
 */
export function convertTablesAndCalcChain(
  workbookPackage: WorkbookPackage,
): void {
  for (const partName of workbookPackage.partsMatching(TABLE_PART_PATTERN)) {
    workbookPackage.writeText(
      partName,
      workbookPackage.requireText(partName).replace(TABLE_FORMULA_PATTERN, ""),
    );
  }
  workbookPackage.removePartAndReferences(CALC_CHAIN_PART, WORKBOOK_PART);
}

export async function convertWorkbookToValues(
  workbookBytes: Uint8Array,
): Promise<Uint8Array> {
  return (await convertWorkbookToValuesWithReport(workbookBytes)).bytes;
}

export async function convertWorkbookToValuesWithReport(
  workbookBytes: Uint8Array,
): Promise<ValuesOnlyConversion> {
  const workbookPackage = await WorkbookPackage.load(workbookBytes);
  const names = worksheetNamesByPart(workbookPackage);
  let formulasConverted = 0;
  const formulasWithoutCachedValues: MissingCachedFormula[] = [];

  for (const partName of workbookPackage.partsMatching(
    WORKSHEET_PART_PATTERN,
  )) {
    const conversion = removeWorksheetFormulas(
      workbookPackage.requireText(partName),
      partName,
    );
    workbookPackage.writeText(partName, conversion.xml);
    formulasConverted += conversion.formulasConverted;
    for (const missing of conversion.formulasWithoutCachedValues) {
      formulasWithoutCachedValues.push({
        ...missing,
        location: `${names.get(partName) ?? partName}!${missing.cell}`,
      });
    }
  }

  convertTablesAndCalcChain(workbookPackage);

  return {
    bytes: await workbookPackage.save(),
    formulasConverted,
    formulasWithoutCachedValues,
  };
}
