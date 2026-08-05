import { copyFile, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";
import { afterEach, describe, expect, it } from "vitest";
import * as XLSX from "xlsx";

import {
  readWorkbookExcelTables,
  splitWorkbookByColumn,
} from "../src/index.js";
import { readExcelTableDefinitions } from "../src/excel-tables.js";
import { preserveWorkbookWithFilteredExcelTable } from "../src/preserve-table-split.js";

const temporaryDirectories: string[] = [];
const structuredTableFixture = fileURLToPath(
  new URL("./fixtures/structured-table.xlsx", import.meta.url),
);

async function temporaryDirectory(): Promise<string> {
  const directory = await mkdtemp(
    path.join(tmpdir(), "consultchimps-full-workbook-split-"),
  );
  temporaryDirectories.push(directory);
  return directory;
}

async function createPreservationWorkbook(filePath: string): Promise<void> {
  const workbook = XLSX.utils.book_new();
  const first = XLSX.utils.aoa_to_sheet([
    ["Entity allocation report"],
    [],
    ["Record", "Entity Name", "Calculated"],
    [1, "DGE", { f: "A4*10", t: "n", v: 10, z: "#,##0.00" }],
    [2, "dge ", { f: "A5*10", t: "n", v: 20, z: "#,##0.00" }],
    [3, "Other", { f: "A6*10", t: "n", v: 30, z: "#,##0.00" }],
    [4, null, 40],
  ]);
  first["!cols"] = [{ wch: 14 }, { wch: 28 }, { wch: 20 }];
  first["!rows"] = [{ hpt: 30 }, {}, { hpt: 22 }];
  first["!merges"] = [XLSX.utils.decode_range("A1:C1")];
  first.A4!.l = { Target: "https://example.com/record/1" };
  first.B4!.c = [{ a: "User", t: "Synthetic test note" }];
  XLSX.utils.book_append_sheet(workbook, first, "Operations");

  const second = XLSX.utils.aoa_to_sheet([
    ["Code", "Amount", "Entity Name"],
    ["A", 5, " DGE"],
    ["B", 6, "Third"],
    ["C", 7, null],
  ]);
  second["!cols"] = [{ wch: 12 }, { wch: 18 }, { wch: 30 }];
  XLSX.utils.book_append_sheet(workbook, second, "Finance");

  const cover = XLSX.utils.aoa_to_sheet([
    ["Cover sheet", { f: "1+1", t: "n" }],
    ["Copied without filtering"],
  ]);
  cover["!merges"] = [XLSX.utils.decode_range("A2:C2")];
  cover["!cols"] = [{ wch: 36 }, { wch: 16 }];
  XLSX.utils.book_append_sheet(workbook, cover, "Cover");
  workbook.Workbook = {
    Sheets: [
      { Hidden: 0, name: "Operations" },
      { Hidden: 0, name: "Finance" },
      { Hidden: 2, name: "Cover" },
    ],
  };

  const archive = await JSZip.loadAsync(
    XLSX.write(workbook, {
      bookType: "xlsx",
      cellStyles: true,
      type: "buffer",
    }),
  );
  const worksheet = await archive
    .file("xl/worksheets/sheet1.xml")!
    .async("text");
  archive.file(
    "xl/worksheets/sheet1.xml",
    worksheet.replace(
      "</worksheet>",
      '<conditionalFormatting sqref="C4:C7"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting><dataValidations count="1"><dataValidation type="whole" sqref="A4:A7"><formula1>1</formula1><formula2>99</formula2></dataValidation></dataValidations></worksheet>',
    ),
  );
  await writeFile(
    filePath,
    await archive.generateAsync({ compression: "DEFLATE", type: "nodebuffer" }),
  );
}

function readWorkbook(filePath: string): Promise<XLSX.WorkBook> {
  return readFile(filePath).then((bytes) =>
    XLSX.read(bytes, { cellStyles: true, type: "buffer" }),
  );
}

function records(
  workbook: XLSX.WorkBook,
  sheet: string,
  headerRow: number,
): Array<Record<string, unknown>> {
  return XLSX.utils.sheet_to_json(workbook.Sheets[sheet]!, {
    defval: null,
    range: headerRow - 1,
  });
}

afterEach(async () => {
  await Promise.all(
    temporaryDirectories
      .splice(0)
      .map((directory) => rm(directory, { force: true, recursive: true })),
  );
});

describe("all-worksheet workbook splitting", () => {
  it("collects normalized values across sheets and preserves the complete workbook", async () => {
    const directory = await temporaryDirectory();
    const input = path.join(directory, "entities.xlsx");
    const output = path.join(directory, "entities");
    await createPreservationWorkbook(input);
    const originalBytes = await readFile(input);
    const originalArchive = await JSZip.loadAsync(originalBytes);
    const originalCoverXml = await originalArchive
      .file("xl/worksheets/sheet3.xml")!
      .async("text");

    const result = await splitWorkbookByColumn({
      column: "Entity Name",
      input,
      outputDirectory: output,
    });

    expect(
      result.artifacts.map((artifact) => path.basename(artifact.path)),
    ).toEqual(["DGE.xlsx", "Other.xlsx", "Third.xlsx"]);
    expect(result.metrics).toMatchObject({
      groups: 3,
      outputFiles: 3,
      sheetsCopiedUnchanged: 1,
      sheetsFiltered: 2,
      skippedRows: 2,
      valuesOnly: 0,
    });
    expect(result.summary).toEqual({
      column: "Entity Name",
      copiedUnchangedSheets: ["Cover"],
      filteredSheets: ["Operations", "Finance"],
      input: path.resolve(input),
      outputDirectory: path.resolve(output),
      valuesOnly: false,
    });
    expect(result.outputs?.[0]).toMatchObject({
      sheets: [
        { deletedRows: 2, retainedRows: 2, sheet: "Operations" },
        { deletedRows: 2, retainedRows: 1, sheet: "Finance" },
      ],
      value: "DGE",
    });

    const dge = await readWorkbook(path.join(output, "DGE.xlsx"));
    expect(dge.SheetNames).toEqual(["Operations", "Finance", "Cover"]);
    expect(records(dge, "Operations", 3)).toEqual([
      { Calculated: 10, "Entity Name": "DGE", Record: 1 },
      { Calculated: 20, "Entity Name": "dge ", Record: 2 },
    ]);
    expect(records(dge, "Finance", 1)).toEqual([
      { Amount: 5, Code: "A", "Entity Name": " DGE" },
    ]);
    expect(dge.Sheets.Operations!.C4?.f).toBe("A4*10");
    expect(dge.Workbook?.Sheets?.[2]?.Hidden).toBe(2);
    expect(dge.Sheets.Cover?.["!merges"]).toEqual([
      XLSX.utils.decode_range("A2:C2"),
    ]);
    expect(dge.Sheets.Operations?.["!cols"]?.[1]?.wch).toBeCloseTo(28, 0);
    expect(dge.Sheets.Operations?.["!rows"]?.[0]?.hpt).toBe(30);

    const dgeArchive = await JSZip.loadAsync(
      await readFile(path.join(output, "DGE.xlsx")),
    );
    expect(
      await dgeArchive.file("xl/worksheets/sheet3.xml")!.async("text"),
    ).toBe(originalCoverXml);
    const operationsXml = await dgeArchive
      .file("xl/worksheets/sheet1.xml")!
      .async("text");
    expect(operationsXml).toContain("conditionalFormatting");
    expect(operationsXml).toContain("dataValidations");
    expect(operationsXml).toContain("hyperlink");
    expect(await readFile(input)).toEqual(originalBytes);
  });

  it("creates safe stable filenames, unifies numeric text, and supports strict matching", async () => {
    const directory = await temporaryDirectory();
    const input = path.join(directory, "names.xlsx");
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet([
        ["Entity Name"],
        ["A/B"],
        ["A:B"],
        [1],
        ["1"],
        ["CON"],
        ["Arabic العربية"],
        ["Trailing. "],
      ]),
      "Data",
    );
    await writeFile(
      input,
      XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }),
    );

    const normalized = await splitWorkbookByColumn({
      column: "Entity Name",
      input,
      outputDirectory: path.join(directory, "normalized"),
    });
    expect(
      normalized.artifacts.map((artifact) => path.basename(artifact.path)),
    ).toEqual([
      "A-B.xlsx",
      "A-B-2.xlsx",
      "1.xlsx",
      "_CON.xlsx",
      "Arabic العربية.xlsx",
      "Trailing.xlsx",
    ]);

    const strict = await splitWorkbookByColumn({
      column: "Entity Name",
      input,
      outputDirectory: path.join(directory, "strict"),
      strict: true,
    });
    expect(strict.metrics.groups).toBe(7);
  });

  it("preflights every normalized destination before writing", async () => {
    const directory = await temporaryDirectory();
    const input = path.join(directory, "entities.xlsx");
    const output = path.join(directory, "entities");
    await createPreservationWorkbook(input);
    await splitWorkbookByColumn({
      column: "Entity Name",
      input,
      outputDirectory: output,
    });

    await expect(
      splitWorkbookByColumn({
        column: "Entity Name",
        input,
        outputDirectory: output,
      }),
    ).rejects.toThrowError(/Output already exists/);
    await expect(
      splitWorkbookByColumn({
        column: "Entity Name",
        input,
        outputDirectory: output,
        overwrite: true,
      }),
    ).resolves.toMatchObject({ metrics: { outputFiles: 3 } });
  });

  it("preserves formulas by default and converts cached values with warnings", async () => {
    const directory = await temporaryDirectory();
    const input = path.join(directory, "entities.xlsx");
    await createPreservationWorkbook(input);

    const result = await splitWorkbookByColumn({
      column: "Entity Name",
      input,
      outputDirectory: path.join(directory, "values"),
      values: true,
    });
    const dge = await readWorkbook(path.join(directory, "values", "DGE.xlsx"));
    expect(dge.Sheets.Operations!.C4?.f).toBeUndefined();
    expect(dge.Sheets.Operations!.C4?.v).toBe(10);
    expect(dge.Sheets.Cover!.B1?.f).toBeUndefined();
    expect(dge.Sheets.Cover!.B1?.v).toBeUndefined();
    expect(result.metrics.formulaCellsConverted).toBeGreaterThan(0);
    expect(result.metrics.formulaCellsWithoutCachedValues).toBe(3);
    expect(result.warnings.join("\n")).toMatch(/Cover!B1/);
    expect(result.warnings.join("\n")).toMatch(
      /recalculate the source workbook/,
    );
  });

  it("updates Excel Table ranges while preserving every worksheet", async () => {
    const directory = await temporaryDirectory();
    const input = path.join(directory, "clients.xlsx");
    await copyFile(structuredTableFixture, input);
    const result = await splitWorkbookByColumn({
      column: "Region",
      input,
      outputDirectory: path.join(directory, "tables"),
    });
    expect(
      result.artifacts.map((artifact) => path.basename(artifact.path)),
    ).toEqual(["North.xlsx", "South.xlsx"]);
    const northTables = await readWorkbookExcelTables(
      path.join(directory, "tables", "North.xlsx"),
    );
    expect(northTables[0]).toMatchObject({
      excelTableRange: "B4:D7",
      rows: [
        { Amount: 10, Client: "A", Region: "North" },
        { Amount: 30, Client: "C", Region: "North" },
      ],
    });
    expect((await readWorkbook(result.artifacts[0]!.path)).SheetNames).toEqual([
      "Cover",
      "Clients",
    ]);
  });

  it("removes complete unmatched table rows, including cells outside the table", async () => {
    const directory = await temporaryDirectory();
    const input = path.join(directory, "clients.xlsx");
    const archive = await JSZip.loadAsync(
      await readFile(structuredTableFixture),
    );
    const sheetPart = archive.file("xl/worksheets/sheet2.xml")!;
    const sheetXml = await sheetPart.async("text");
    archive.file(
      "xl/worksheets/sheet2.xml",
      sheetXml.replace(
        /(<row[^>]*\br="6"[^>]*>)/u,
        '$1<c r="G6" t="inlineStr"><is><t>South confidential note</t></is></c>',
      ),
    );
    await writeFile(
      input,
      await archive.generateAsync({
        compression: "DEFLATE",
        type: "nodebuffer",
      }),
    );

    const result = await splitWorkbookByColumn({
      column: "Region",
      input,
      outputDirectory: path.join(directory, "split"),
    });
    const northArchive = await JSZip.loadAsync(
      await readFile(
        result.artifacts.find((artifact) =>
          artifact.path.endsWith("North.xlsx"),
        )!.path,
      ),
    );
    const northXml = await northArchive
      .file("xl/worksheets/sheet2.xml")!
      .async("text");
    expect(northXml).not.toContain("South confidential note");
  });

  it("supports a table with no rows for a group", async () => {
    const definitions = await readExcelTableDefinitions(
      await readFile(structuredTableFixture),
    );
    const definition = definitions[0];
    expect(definition).toBeDefined();
    const output = await preserveWorkbookWithFilteredExcelTable(
      await readFile(structuredTableFixture),
      { definition: definition!, sourceRows: [], wholeRows: true },
    );
    const outputArchive = await JSZip.loadAsync(output);
    const tableXml = await outputArchive
      .file(definition!.tablePart)!
      .async("text");
    expect(tableXml).toContain('ref="B4:D5"');
  });

  it("deletes complete table rows when preserving a selected group", async () => {
    const definitions = await readExcelTableDefinitions(
      await readFile(structuredTableFixture),
    );
    const definition = definitions[0];
    expect(definition).toBeDefined();
    const output = await preserveWorkbookWithFilteredExcelTable(
      await readFile(structuredTableFixture),
      {
        definition: definition!,
        sourceRows: [5],
        wholeRows: true,
        values: true,
      },
    );
    const outputArchive = await JSZip.loadAsync(output);
    const worksheetXml = await outputArchive
      .file(definition!.worksheetPart)!
      .async("text");
    expect(worksheetXml).toContain('<x:row r="5">');
    // The totals row is retained and compacted directly after the selected data.
    expect(worksheetXml).toContain('<x:row r="6">');
    expect(worksheetXml).not.toContain('<x:row r="7">');
    expect(worksheetXml).not.toContain('<x:row r="8">');

    // Exercise the legacy cell-only mode as a compatibility guard.
    const cellOnlyOutput = await preserveWorkbookWithFilteredExcelTable(
      await readFile(structuredTableFixture),
      { definition: definition!, sourceRows: [5], wholeRows: false },
    );
    const cellOnlyArchive = await JSZip.loadAsync(cellOnlyOutput);
    const cellOnlyXml = await cellOnlyArchive
      .file(definition!.worksheetPart)!
      .async("text");
    expect(cellOnlyXml).toContain('<x:row r="6">');
  });

  it("retains macro package content and the .xlsm extension", async () => {
    const directory = await temporaryDirectory();
    const xlsx = path.join(directory, "source.xlsx");
    const input = path.join(directory, "source.xlsm");
    await createPreservationWorkbook(xlsx);
    const archive = await JSZip.loadAsync(await readFile(xlsx));
    archive.file("xl/vbaProject.bin", Buffer.from([0, 1, 2, 3, 4]));
    const contentTypes = await archive
      .file("[Content_Types].xml")!
      .async("text");
    archive.file(
      "[Content_Types].xml",
      contentTypes.replace(
        "</Types>",
        '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>',
      ),
    );
    const relationships = await archive
      .file("xl/_rels/workbook.xml.rels")!
      .async("text");
    archive.file(
      "xl/_rels/workbook.xml.rels",
      relationships.replace(
        "</Relationships>",
        '<Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>',
      ),
    );
    await writeFile(input, await archive.generateAsync({ type: "nodebuffer" }));

    const result = await splitWorkbookByColumn({
      column: "Entity Name",
      input,
      outputDirectory: path.join(directory, "macros"),
    });
    expect(
      result.artifacts.every((artifact) => artifact.path.endsWith(".xlsm")),
    ).toBe(true);
    const outputArchive = await JSZip.loadAsync(
      await readFile(result.artifacts[0]!.path),
    );
    expect(
      await outputArchive.file("xl/vbaProject.bin")!.async("nodebuffer"),
    ).toEqual(Buffer.from([0, 1, 2, 3, 4]));
  });

  it("reports missing columns, blank columns, and unsupported input types", async () => {
    const directory = await temporaryDirectory();
    const input = path.join(directory, "blank.xlsx");
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet([
        ["Entity Name", "Value"],
        [null, 1],
      ]),
      "Data",
    );
    await writeFile(
      input,
      XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }),
    );

    await expect(
      splitWorkbookByColumn({
        column: "Missing",
        input,
        outputDirectory: path.join(directory, "missing"),
      }),
    ).rejects.toMatchObject({ code: "XLSX_SPLIT_COLUMN_NOT_FOUND" });
    await expect(
      splitWorkbookByColumn({
        column: "Entity Name",
        input,
        outputDirectory: path.join(directory, "empty"),
      }),
    ).rejects.toMatchObject({ code: "XLSX_SPLIT_NO_GROUPS" });
    await expect(
      splitWorkbookByColumn({
        column: "Entity Name",
        input: path.join(directory, "book.xls"),
        outputDirectory: path.join(directory, "unsupported"),
      }),
    ).rejects.toMatchObject({ code: "XLSX_SPLIT_UNSUPPORTED_FILE" });
  });
});
