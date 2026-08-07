import { WorkbookPackage } from "./package/index.js";

const CELL_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?c\b[^>]*?(?:\/\s*>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?c\s*>)/gu;
const CELL_FORMULA_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?f\b[^>]*(?:\/\s*>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?f\s*>)/gu;
const TABLE_FORMULA_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?(?:calculatedColumnFormula|totalsRowFormula)\b[^>]*(?:\/\s*>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?(?:calculatedColumnFormula|totalsRowFormula)\s*>)/gu;
const CACHED_VALUE_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?(?:v|is)\b[^>]*(?:\/\s*>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?(?:v|is)\s*>)/u;
const CELL_REFERENCE_PATTERN = /\br=(?:"([^"]+)"|'([^']+)')/u;
const WORKSHEET_PART_PATTERN = /^xl\/worksheets\/[^/]+\.xml$/iu;
const TABLE_PART_PATTERN = /^xl\/tables\/[^/]+\.xml$/iu;
const CALC_CHAIN_PART = "xl/calcChain.xml";
const WORKBOOK_PART = "xl/workbook.xml";

export interface MissingCachedFormula {
  cell: string;
  worksheetPart: string;
}

export interface ValuesOnlyConversion {
  bytes: Uint8Array;
  formulasConverted: number;
  formulasWithoutCachedValues: MissingCachedFormula[];
}

function removeWorksheetFormulas(
  worksheetXml: string,
  worksheetPart: string,
): {
  formulasConverted: number;
  formulasWithoutCachedValues: MissingCachedFormula[];
  xml: string;
} {
  let formulasConverted = 0;
  const formulasWithoutCachedValues: MissingCachedFormula[] = [];
  const xml = worksheetXml.replace(CELL_PATTERN, (cellXml) => {
    if (!CELL_FORMULA_PATTERN.test(cellXml)) {
      return cellXml;
    }

    CELL_FORMULA_PATTERN.lastIndex = 0;
    formulasConverted += 1;
    if (!CACHED_VALUE_PATTERN.test(cellXml)) {
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
export async function convertWorkbookToValues(
  workbookBytes: Uint8Array,
): Promise<Uint8Array> {
  return (await convertWorkbookToValuesWithReport(workbookBytes)).bytes;
}

export async function convertWorkbookToValuesWithReport(
  workbookBytes: Uint8Array,
): Promise<ValuesOnlyConversion> {
  const workbookPackage = await WorkbookPackage.load(workbookBytes);
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
    formulasWithoutCachedValues.push(...conversion.formulasWithoutCachedValues);
  }

  for (const partName of workbookPackage.partsMatching(TABLE_PART_PATTERN)) {
    workbookPackage.writeText(
      partName,
      workbookPackage.requireText(partName).replace(TABLE_FORMULA_PATTERN, ""),
    );
  }

  // A workbook with no formulas left has nothing to calculate.
  workbookPackage.removePartAndReferences(CALC_CHAIN_PART, WORKBOOK_PART);

  return {
    bytes: await workbookPackage.save(),
    formulasConverted,
    formulasWithoutCachedValues,
  };
}
