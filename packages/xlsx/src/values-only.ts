import JSZip from "jszip";

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

function removeWorksheetFormulas(worksheetXml: string): string {
  return worksheetXml.replace(CELL_PATTERN, (cellXml) =>
    cellXml.replace(CELL_FORMULA_PATTERN, ""),
  );
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
  workbookBytes: Buffer,
): Promise<Buffer> {
  const archive = await JSZip.loadAsync(workbookBytes);
  const worksheetParts = Object.keys(archive.files).filter((partName) =>
    /^xl\/worksheets\/[^/]+\.xml$/iu.test(partName),
  );
  const tableParts = Object.keys(archive.files).filter((partName) =>
    /^xl\/tables\/[^/]+\.xml$/iu.test(partName),
  );

  await Promise.all(
    worksheetParts.map(async (partName) => {
      const entry = archive.file(partName);
      if (entry) {
        archive.file(
          partName,
          removeWorksheetFormulas(await entry.async("text")),
        );
      }
    }),
  );
  await Promise.all(
    tableParts.map(async (partName) => {
      const entry = archive.file(partName);
      if (entry) {
        archive.file(
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
      archive.file(
        partName,
        removeCalculationChainReferences(await entry.async("text")),
      );
    }
  }

  return archive.generateAsync({
    compression: "DEFLATE",
    type: "nodebuffer",
  });
}
