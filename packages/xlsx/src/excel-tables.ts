/**
 * Readers for the workbook parts that describe worksheets and Excel Tables.
 * Package access goes through L0; nothing here opens an archive itself.
 */
import {
  forEachOpenTag,
  tagAttribute,
  WorkbookPackage,
} from "./package/index.js";

const TABLE_RELATIONSHIP_SUFFIX = "/table";
const WORKSHEET_RELATIONSHIP_SUFFIX = "/worksheet";
const WORKBOOK_PART = "xl/workbook.xml";

interface WorkbookSheet {
  name: string;
  relationshipId: string;
  sheetId: number;
  state: string | undefined;
}

/** A worksheet as the workbook part declares it, with its package part path. */
export interface WorkbookSheetEntry {
  name: string;
  /** The workbook's own sheet id, which calculation-chain entries index by. */
  sheetId: number;
  state: string | undefined;
  worksheetPart: string;
}

export type WorkbookWorksheetPart = WorkbookSheetEntry;

export interface ExcelTableDefinition {
  columns: string[];
  headerRow: boolean;
  name: string;
  range: string;
  sheet: string;
  tablePart: string;
  totalsRow: boolean;
  worksheetPart: string;
}

function parseWorkbookSheets(xml: string, fileName: string): WorkbookSheet[] {
  const sheets: WorkbookSheet[] = [];

  forEachOpenTag(xml, fileName, (tag) => {
    if (tag.local !== "sheet") {
      return;
    }

    const name = tagAttribute(tag, "name");
    const relationshipId = tagAttribute(tag, "id");
    if (name && relationshipId) {
      sheets.push({
        name,
        relationshipId,
        sheetId: Number(tagAttribute(tag, "sheetId") ?? sheets.length + 1),
        state: tagAttribute(tag, "state"),
      });
    }
  });

  return sheets;
}

function workbookSheets(workbookPackage: WorkbookPackage): WorkbookSheet[] {
  return parseWorkbookSheets(
    workbookPackage.requireText(WORKBOOK_PART),
    WORKBOOK_PART,
  );
}

/** Every worksheet the workbook declares, in workbook order. */
export function readWorkbookSheetsFrom(
  workbookPackage: WorkbookPackage,
): WorkbookSheetEntry[] {
  const relationships = new Map(
    workbookPackage
      .relationshipsOf(WORKBOOK_PART)
      .map((relationship) => [relationship.id, relationship] as const),
  );

  return workbookSheets(workbookPackage).flatMap((sheet) => {
    const relationship = relationships.get(sheet.relationshipId);
    if (!relationship?.type.endsWith(WORKSHEET_RELATIONSHIP_SUFFIX)) {
      return [];
    }
    return [
      {
        name: sheet.name,
        sheetId: sheet.sheetId,
        state: sheet.state,
        worksheetPart: workbookPackage.resolvePart(
          WORKBOOK_PART,
          relationship.target,
        ),
      },
    ];
  });
}

function parseTableDefinition(
  xml: string,
  fileName: string,
  sheet: string,
  tablePart: string,
  worksheetPart: string,
): ExcelTableDefinition {
  let definition:
    | Omit<
        ExcelTableDefinition,
        "columns" | "sheet" | "tablePart" | "worksheetPart"
      >
    | undefined;
  const columns: string[] = [];

  forEachOpenTag(xml, fileName, (tag) => {
    if (tag.local === "table") {
      const name =
        tagAttribute(tag, "displayName") ?? tagAttribute(tag, "name");
      const range = tagAttribute(tag, "ref");
      if (name && range) {
        definition = {
          headerRow: tagAttribute(tag, "headerRowCount") !== "0",
          name,
          range,
          totalsRow: Number(tagAttribute(tag, "totalsRowCount") ?? "0") > 0,
        };
      }
      return;
    }

    if (tag.local === "tableColumn") {
      const name = tagAttribute(tag, "name");
      if (name !== undefined) {
        columns.push(name);
      }
    }
  });

  if (!definition) {
    throw new Error(
      `No valid Excel Table definition was found in ${fileName}.`,
    );
  }
  if (columns.length === 0) {
    throw new Error(`Excel Table "${definition.name}" has no columns.`);
  }

  return {
    ...definition,
    columns,
    sheet,
    tablePart,
    worksheetPart,
  };
}

/** Every Excel Table the package declares, in worksheet order. */
export function readExcelTableDefinitionsFrom(
  workbookPackage: WorkbookPackage,
): ExcelTableDefinition[] {
  const definitions: ExcelTableDefinition[] = [];

  for (const sheet of readWorkbookSheetsFrom(workbookPackage)) {
    const tableRelationships = workbookPackage
      .relationshipsOf(sheet.worksheetPart)
      .filter((relationship) =>
        relationship.type.endsWith(TABLE_RELATIONSHIP_SUFFIX),
      );

    for (const tableRelationship of tableRelationships) {
      const tablePart = workbookPackage.resolvePart(
        sheet.worksheetPart,
        tableRelationship.target,
      );
      definitions.push(
        parseTableDefinition(
          workbookPackage.requireText(tablePart),
          tablePart,
          sheet.name,
          tablePart,
          sheet.worksheetPart,
        ),
      );
    }
  }

  return definitions;
}

export async function readWorkbookWorksheetParts(
  workbookBytes: Uint8Array,
): Promise<WorkbookSheetEntry[]> {
  return readWorkbookSheetsFrom(await WorkbookPackage.load(workbookBytes));
}

export async function readExcelTableDefinitions(
  workbookBytes: Uint8Array,
): Promise<ExcelTableDefinition[]> {
  return readExcelTableDefinitionsFrom(
    await WorkbookPackage.load(workbookBytes),
  );
}
