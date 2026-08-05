import JSZip from "jszip";

import { generatePackageBytes, replacePackagePart } from "./package-zip.js";

const CELL_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?c\b[^>]*?(?:\/\s*>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?c\s*>)/gu;
const CELL_FORMULA_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?f\b[^>]*(?:\/\s*>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?f\s*>)/gu;
const TABLE_FORMULA_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?(?:calculatedColumnFormula|totalsRowFormula)\b[^>]*(?:\/\s*>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?(?:calculatedColumnFormula|totalsRowFormula)\s*>)/gu;
const CALC_CHAIN_RELATIONSHIP_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?Relationship\b(?=[^>]*\bType=(?:"[^"]*\/calcChain"|'[^']*\/calcChain'))[^>]*\/\s*>/gu;
const CALC_CHAIN_CONTENT_TYPE_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?Override\b(?=[^>]*\bPartName=(?:"\/xl\/calcChain\.xml"|'\/xl\/calcChain\.xml'))[^>]*\/\s*>/gu;
const CACHED_VALUE_PATTERN =
  /<(?:[A-Za-z_][\w.-]*:)?(?:v|is)\b[^>]*(?:\/\s*>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?(?:v|is)\s*>)/u;
const CELL_REFERENCE_PATTERN = /\br=(?:"([^"]+)"|'([^']+)')/u;

export interface MissingCachedFormula {
  cell: string;
  worksheetPart: string;
}

export interface ValuesOnlyConversion {
  bytes: Buffer;
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

function removeCalculationChainReferences(xml: string): string {
  return xml
    .replace(CALC_CHAIN_RELATIONSHIP_PATTERN, "")
    .replace(CALC_CHAIN_CONTENT_TYPE_PATTERN, "");
}

/**
 * Replace worksheet formulas with their cached values without rebuilding any
 * cells. Editing the OOXML parts directly preserves cell styles, number
 * formats, row heights, column widths, tables, and the rest of the workbook
 * package byte-for-byte apart from formula and calculation-chain metadata.
 */
export async function convertWorkbookToValues(
  workbookBytes: Uint8Array,
): Promise<Buffer> {
  return (await convertWorkbookToValuesWithReport(workbookBytes)).bytes;
}

export async function convertWorkbookToValuesWithReport(
  workbookBytes: Uint8Array,
): Promise<ValuesOnlyConversion> {
  const archive = await JSZip.loadAsync(workbookBytes);
  const worksheetParts = Object.keys(archive.files).filter((partName) =>
    /^xl\/worksheets\/[^/]+\.xml$/iu.test(partName),
  );
  const tableParts = Object.keys(archive.files).filter((partName) =>
    /^xl\/tables\/[^/]+\.xml$/iu.test(partName),
  );

  const worksheetConversions = await Promise.all(
    worksheetParts.map(async (partName) => {
      const entry = archive.file(partName);
      if (entry) {
        const conversion = removeWorksheetFormulas(
          await entry.async("text"),
          partName,
        );
        replacePackagePart(archive, partName, conversion.xml);
        return conversion;
      }
      return undefined;
    }),
  );
  await Promise.all(
    tableParts.map(async (partName) => {
      const entry = archive.file(partName);
      if (entry) {
        replacePackagePart(
          archive,
          partName,
          (await entry.async("text")).replace(TABLE_FORMULA_PATTERN, ""),
        );
      }
    }),
  );

  archive.remove("xl/calcChain.xml");
  for (const partName of [
    "xl/_rels/workbook.xml.rels",
    "[Content_Types].xml",
  ]) {
    const entry = archive.file(partName);
    if (entry) {
      replacePackagePart(
        archive,
        partName,
        removeCalculationChainReferences(await entry.async("text")),
      );
    }
  }

  return {
    bytes: Buffer.from(await generatePackageBytes(archive)),
    formulasConverted: worksheetConversions.reduce(
      (total, conversion) => total + (conversion?.formulasConverted ?? 0),
      0,
    ),
    formulasWithoutCachedValues: worksheetConversions.flatMap(
      (conversion) => conversion?.formulasWithoutCachedValues ?? [],
    ),
  };
}
