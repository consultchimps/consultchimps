import {
  copyFile,
  mkdtemp,
  readFile,
  rm,
  stat,
  writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";
import { afterEach, describe, expect, it } from "vitest";

import {
  readWorkbookExcelTables,
  splitWorkbookByColumn,
} from "../src/index.js";
import { readExcelTableDefinitions } from "../src/excel-tables.js";
import { preserveWorkbookWithFilteredExcelTable } from "../src/preserve-table-split.js";
import {
  mergedCellReferences,
  worksheetCellFormula,
  worksheetCellValue,
} from "./corpus/fixtures.js";
import { sheetNamesOf, sheetRows } from "./support/read-workbook.js";
import {
  buildSheetFixture,
  buildWorkbookFixture,
} from "./support/workbook-fixture.js";

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

const COMMENTS_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";

async function createPreservationWorkbook(filePath: string): Promise<void> {
  const money = (formula: string, value: number) => ({
    formula,
    value,
    format: "#,##0.00",
  });
  const archive = await JSZip.loadAsync(
    await buildWorkbookFixture({
      sheets: [
        {
          name: "Operations",
          rows: [
            ["Entity allocation report"],
            [],
            ["Record", "Entity Name", "Calculated"],
            [1, "Alpha", money("A4*10", 10)],
            [2, "alpha ", money("A5*10", 20)],
            [3, "Other", money("A6*10", 30)],
            [4, null, 40],
          ],
          merges: ["A1:C1"],
          widths: [14, 28, 20],
        },
        {
          name: "Finance",
          rows: [
            ["Code", "Amount", "Entity Name"],
            ["A", 5, " Alpha"],
            ["B", 6, "Third"],
            ["C", 7, null],
          ],
          widths: [12, 18, 30],
        },
        {
          name: "Cover",
          rows: [
            ["Cover sheet", { formula: "1+1" }],
            ["Copied without filtering"],
          ],
          merges: ["A2:C2"],
          widths: [36, 16],
          state: "veryHidden",
        },
      ],
    }),
  );
  // Row heights, a hyperlink, a comment, conditional formatting and data
  // validation, which the builder does not write, are added by hand.
  const worksheet = await archive
    .file("xl/worksheets/sheet1.xml")!
    .async("text");
  archive.file(
    "xl/worksheets/sheet1.xml",
    worksheet
      .replace('<row r="1">', '<row r="1" ht="30" customHeight="1">')
      .replace('<row r="3">', '<row r="3" ht="22" customHeight="1">')
      .replace(
        "</worksheet>",
        '<conditionalFormatting sqref="C4:C7"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting><dataValidations count="1"><dataValidation type="whole" sqref="A4:A7"><formula1>1</formula1><formula2>99</formula2></dataValidation></dataValidations><hyperlinks><hyperlink ref="A4" r:id="rId1"/></hyperlinks></worksheet>',
      ),
  );
  archive.file(
    "xl/worksheets/_rels/sheet1.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/record/1" TargetMode="External"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/></Relationships>`,
  );
  archive.file(
    "xl/comments1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>User</author></authors><commentList><comment ref="B4" authorId="0"><text><t>Synthetic test note</t></text></comment></commentList></comments>`,
  );
  const contentTypes = await archive.file("[Content_Types].xml")!.async("text");
  archive.file(
    "[Content_Types].xml",
    contentTypes.replace(
      "</Types>",
      `<Override PartName="/xl/comments1.xml" ContentType="${COMMENTS_CONTENT_TYPE}"/></Types>`,
    ),
  );
  await writeFile(
    filePath,
    await archive.generateAsync({ compression: "DEFLATE", type: "nodebuffer" }),
  );
}

async function packagePart(filePath: string, part: string): Promise<string> {
  const archive = await JSZip.loadAsync(await readFile(filePath));
  return archive.file(part)!.async("text");
}

/** Records under a header row, empty cells as null, blank rows skipped. */
async function records(
  filePath: string,
  sheet: string,
  headerRow: number,
): Promise<Array<Record<string, unknown>>> {
  const [header = [], ...body] = (
    await sheetRows(await readFile(filePath), sheet)
  ).slice(headerRow - 1);
  return body
    .filter((row) => row.some((value) => value !== null))
    .map((row) =>
      Object.fromEntries(
        header.map((key, index) => [String(key), row[index] ?? null]),
      ),
    );
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
    ).toEqual(["Alpha.xlsx", "Other.xlsx", "Third.xlsx"]);
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
      value: "Alpha",
    });

    const alphaPath = path.join(output, "Alpha.xlsx");
    expect(await sheetNamesOf(await readFile(alphaPath))).toEqual([
      "Operations",
      "Finance",
      "Cover",
    ]);
    expect(await records(alphaPath, "Operations", 3)).toEqual([
      { Calculated: 10, "Entity Name": "Alpha", Record: 1 },
      { Calculated: 20, "Entity Name": "alpha ", Record: 2 },
    ]);
    expect(await records(alphaPath, "Finance", 1)).toEqual([
      { Amount: 5, Code: "A", "Entity Name": " Alpha" },
    ]);
    const dgeOperations = await packagePart(
      alphaPath,
      "xl/worksheets/sheet1.xml",
    );
    expect(worksheetCellFormula(dgeOperations, "C4")).toBe("A4*10");
    expect(await packagePart(alphaPath, "xl/workbook.xml")).toMatch(
      /<sheet\b[^>]*\bname="Cover"[^>]*\bstate="veryHidden"/u,
    );
    expect(
      mergedCellReferences(
        await packagePart(alphaPath, "xl/worksheets/sheet3.xml"),
      ),
    ).toEqual(["A2:C2"]);
    expect(
      Number(
        /<col\b[^>]*\bmin="2"[^>]*\bwidth="([\d.]+)"/u.exec(dgeOperations)?.[1],
      ),
    ).toBeCloseTo(28, 0);
    expect(dgeOperations).toMatch(/<row\b[^>]*\br="1"[^>]*\bht="30"/u);

    const otherArchive = await JSZip.loadAsync(
      await readFile(path.join(output, "Other.xlsx")),
    );
    const otherXml = await otherArchive
      .file("xl/worksheets/sheet1.xml")!
      .async("text");
    expect(otherXml).toMatch(/<row\b[^>]*\br="4"/u);
    expect(otherXml).not.toMatch(/<row\b[^>]*\br="6"/u);

    const dgeArchive = await JSZip.loadAsync(
      await readFile(path.join(output, "Alpha.xlsx")),
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
    await writeFile(
      input,
      await buildSheetFixture("Data", [
        ["Entity Name"],
        ["A/B"],
        ["A:B"],
        [1],
        ["1"],
        ["CON"],
        ["Arabic العربية"],
        ["Trailing. "],
      ]),
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
    const alphaPath = path.join(directory, "values", "Alpha.xlsx");
    const operations = await packagePart(alphaPath, "xl/worksheets/sheet1.xml");
    expect(worksheetCellFormula(operations, "C4")).toBeUndefined();
    expect(worksheetCellValue(operations, "C4")).toBe("10");
    const cover = await packagePart(alphaPath, "xl/worksheets/sheet3.xml");
    expect(worksheetCellFormula(cover, "B1")).toBeUndefined();
    expect(worksheetCellValue(cover, "B1")).toBeUndefined();
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
    expect(
      await sheetNamesOf(await readFile(result.artifacts[0]!.path)),
    ).toEqual(["Cover", "Clients"]);
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
      { definition: definition!, sourceRows: [] },
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
      { definition: definition!, sourceRows: [5], values: true },
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
    // The main workbook part is declared macro-enabled as well as the VBA
    // project being added: a package that carries macros while still calling
    // itself an ordinary workbook is a contradiction the split refuses, so a
    // fixture standing in for a real .xlsm has to declare both.
    const declared = contentTypes
      .replace(
        "</Types>",
        '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>',
      )
      .replace(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
      );
    expect(declared).toContain("macroEnabled.main+xml");
    archive.file("[Content_Types].xml", declared);
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

  it("refuses a workbook whose package contradicts its name", async () => {
    const directory = await temporaryDirectory();
    const xlsx = path.join(directory, "source.xlsx");
    const renamed = path.join(directory, "renamed.xlsm");
    await createPreservationWorkbook(xlsx);
    // Renaming an ordinary workbook is all it takes to reach this: nothing in
    // the package changed, so every output would be advertised as
    // macro-enabled while still declaring an ordinary workbook.
    await writeFile(renamed, await readFile(xlsx));

    await expect(
      splitWorkbookByColumn({
        column: "Entity Name",
        input: renamed,
        outputDirectory: path.join(directory, "out"),
      }),
    ).rejects.toMatchObject({
      code: "XLSX_SPLIT_PACKAGE_TYPE_MISMATCH",
      details: { declaredExtension: ".xlsx", nameExtension: ".xlsm" },
    });
    // The refusal comes before any destination is created.
    await expect(stat(path.join(directory, "out"))).rejects.toMatchObject({
      code: "ENOENT",
    });
  });

  it("reports missing columns, blank columns, and unsupported input types", async () => {
    const directory = await temporaryDirectory();
    const input = path.join(directory, "blank.xlsx");
    await writeFile(
      input,
      await buildSheetFixture("Data", [
        ["Entity Name", "Value"],
        [null, 1],
      ]),
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
