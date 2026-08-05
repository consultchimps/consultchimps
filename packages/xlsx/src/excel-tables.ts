import JSZip from "jszip";
import { SaxesParser, type SaxesTagNS } from "saxes";

import {
  joinPackagePath,
  packagePartDirectory,
  packagePartName,
  normalizePackagePath,
} from "./package-paths.js";

const TABLE_RELATIONSHIP_SUFFIX = "/table";
const WORKSHEET_RELATIONSHIP_SUFFIX = "/worksheet";

interface PackageRelationship {
  id: string;
  target: string;
  type: string;
}

interface WorkbookSheet {
  name: string;
  relationshipId: string;
  state: string | undefined;
}

export interface WorkbookWorksheetPart {
  name: string;
  state: string | undefined;
  worksheetPart: string;
}

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

function attribute(tag: SaxesTagNS, localName: string): string | undefined {
  const requestedName = localName.toLocaleLowerCase();
  return Object.values(tag.attributes).find(
    (candidate) => candidate.local.toLocaleLowerCase() === requestedName,
  )?.value;
}

function parseXml(
  xml: string,
  fileName: string,
  onOpenTag: (tag: SaxesTagNS) => void,
): void {
  const parser = new SaxesParser({
    fileName,
    position: true,
    xmlns: true,
  } as const);
  parser.on("doctype", () => {
    throw new Error(`DOCTYPE declarations are not allowed in ${fileName}.`);
  });
  parser.on("error", (error) => {
    throw error;
  });
  parser.on("opentag", onOpenTag);
  parser.write(xml).close();
}

function parseRelationships(
  xml: string,
  fileName: string,
): PackageRelationship[] {
  const relationships: PackageRelationship[] = [];

  parseXml(xml, fileName, (tag) => {
    if (tag.local !== "Relationship") {
      return;
    }

    const id = attribute(tag, "Id");
    const target = attribute(tag, "Target");
    const type = attribute(tag, "Type");
    if (id && target && type) {
      relationships.push({ id, target, type });
    }
  });

  return relationships;
}

function parseWorkbookSheets(xml: string, fileName: string): WorkbookSheet[] {
  const sheets: WorkbookSheet[] = [];

  parseXml(xml, fileName, (tag) => {
    if (tag.local !== "sheet") {
      return;
    }

    const name = attribute(tag, "name");
    const relationshipId = attribute(tag, "id");
    if (name && relationshipId) {
      sheets.push({
        name,
        relationshipId,
        state: attribute(tag, "state"),
      });
    }
  });

  return sheets;
}

export async function readWorkbookWorksheetParts(
  workbookBytes: Buffer,
): Promise<WorkbookWorksheetPart[]> {
  const archive = await JSZip.loadAsync(workbookBytes);
  const workbookPart = "xl/workbook.xml";
  const sheets = parseWorkbookSheets(
    await requiredText(archive, workbookPart),
    workbookPart,
  );
  const workbookRelationships = new Map(
    (await optionalRelationships(archive, workbookPart)).map(
      (relationship) => [relationship.id, relationship] as const,
    ),
  );

  return sheets.flatMap((sheet) => {
    const relationship = workbookRelationships.get(sheet.relationshipId);
    if (!relationship?.type.endsWith(WORKSHEET_RELATIONSHIP_SUFFIX)) {
      return [];
    }

    return [
      {
        name: sheet.name,
        state: sheet.state,
        worksheetPart: resolvePartPath(workbookPart, relationship.target),
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

  parseXml(xml, fileName, (tag) => {
    if (tag.local === "table") {
      const name = attribute(tag, "displayName") ?? attribute(tag, "name");
      const range = attribute(tag, "ref");
      if (name && range) {
        definition = {
          headerRow: attribute(tag, "headerRowCount") !== "0",
          name,
          range,
          totalsRow: Number(attribute(tag, "totalsRowCount") ?? "0") > 0,
        };
      }
      return;
    }

    if (tag.local === "tableColumn") {
      const name = attribute(tag, "name");
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

function relationshipsPath(partPath: string): string {
  return joinPackagePath(
    packagePartDirectory(partPath),
    "_rels",
    `${packagePartName(partPath)}.rels`,
  );
}

function resolvePartPath(sourcePart: string, target: string): string {
  const candidate = target.startsWith("/")
    ? target.slice(1)
    : joinPackagePath(packagePartDirectory(sourcePart), target);
  const normalized = normalizePackagePath(candidate);

  if (
    normalized === ".." ||
    normalized.startsWith("../") ||
    normalized.startsWith("/")
  ) {
    throw new Error(
      `Relationship target escapes the workbook package: ${target}`,
    );
  }

  return normalized;
}

async function requiredText(archive: JSZip, partPath: string): Promise<string> {
  const entry = archive.file(partPath);
  if (!entry) {
    throw new Error(`Workbook package part is missing: ${partPath}`);
  }
  return entry.async("text");
}

async function optionalRelationships(
  archive: JSZip,
  sourcePart: string,
): Promise<PackageRelationship[]> {
  const partPath = relationshipsPath(sourcePart);
  const entry = archive.file(partPath);
  if (!entry) {
    return [];
  }
  return parseRelationships(await entry.async("text"), partPath);
}

export async function readExcelTableDefinitions(
  workbookBytes: Uint8Array,
): Promise<ExcelTableDefinition[]> {
  const archive = await JSZip.loadAsync(workbookBytes);
  const workbookPart = "xl/workbook.xml";
  const sheets = parseWorkbookSheets(
    await requiredText(archive, workbookPart),
    workbookPart,
  );
  const workbookRelationships = new Map(
    (await optionalRelationships(archive, workbookPart)).map(
      (relationship) => [relationship.id, relationship] as const,
    ),
  );
  const definitions: ExcelTableDefinition[] = [];

  for (const sheet of sheets) {
    const sheetRelationship = workbookRelationships.get(sheet.relationshipId);
    if (
      !sheetRelationship ||
      !sheetRelationship.type.endsWith(WORKSHEET_RELATIONSHIP_SUFFIX)
    ) {
      continue;
    }

    const sheetPart = resolvePartPath(workbookPart, sheetRelationship.target);
    const tableRelationships = (
      await optionalRelationships(archive, sheetPart)
    ).filter((relationship) =>
      relationship.type.endsWith(TABLE_RELATIONSHIP_SUFFIX),
    );

    for (const tableRelationship of tableRelationships) {
      const tablePart = resolvePartPath(sheetPart, tableRelationship.target);
      definitions.push(
        parseTableDefinition(
          await requiredText(archive, tablePart),
          tablePart,
          sheet.name,
          tablePart,
          sheetPart,
        ),
      );
    }
  }

  return definitions;
}
