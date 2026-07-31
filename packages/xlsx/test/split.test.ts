import {
  copyFile,
  mkdir,
  mkdtemp,
  readFile,
  readdir,
  rm,
  writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";
import { afterEach, describe, expect, it } from "vitest";
import * as XLSX from "xlsx";

import {
  planSplitWorkbookByColumn,
  readWorkbookExcelTables,
  splitWorkbookByColumn,
} from "../src/index.js";

const temporaryDirectories: string[] = [];
const structuredTableFixture = fileURLToPath(
  new URL("./fixtures/structured-table.xlsx", import.meta.url),
);

async function createTemporaryDirectory(): Promise<string> {
  const directory = await mkdtemp(
    path.join(tmpdir(), "consultchimps-xlsx-split-"),
  );
  temporaryDirectories.push(directory);
  return directory;
}

async function createWorkbook(
  filePath: string,
  sheets: Array<[string, Array<Array<boolean | null | number | string>>]>,
): Promise<void> {
  const workbook = XLSX.utils.book_new();
  for (const [sheetName, rows] of sheets) {
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet(rows),
      sheetName,
    );
  }
  await writeFile(
    filePath,
    XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }),
  );
}

async function readRows(
  filePath: string,
  sheetName: string,
): Promise<Array<Record<string, unknown>>> {
  const workbook = XLSX.read(await readFile(filePath), { type: "buffer" });
  return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]!, {
    defval: null,
    raw: true,
  });
}

async function copyWorkbookWithExcelTable(
  filePath: string,
  hidden = false,
): Promise<void> {
  await copyFile(structuredTableFixture, filePath);
  if (!hidden) {
    return;
  }

  const archive = await JSZip.loadAsync(await readFile(filePath));
  const workbookEntry = archive.file("xl/workbook.xml");
  if (!workbookEntry) {
    throw new Error("The structured-table fixture has no workbook part.");
  }
  const workbookXml = await workbookEntry.async("text");
  const hiddenWorkbookXml = workbookXml.replace(
    'name="Clients" sheetId="2"',
    'name="Clients" state="hidden" sheetId="2"',
  );
  if (hiddenWorkbookXml === workbookXml) {
    throw new Error("The Clients worksheet was not found in the fixture.");
  }
  archive.file("xl/workbook.xml", hiddenWorkbookXml);
  await writeFile(
    filePath,
    await archive.generateAsync({ compression: "DEFLATE", type: "nodebuffer" }),
  );
}

afterEach(async () => {
  await Promise.all(
    temporaryDirectories
      .splice(0)
      .map((directory) => rm(directory, { force: true, recursive: true })),
  );
});

describe("splitWorkbookByColumn", () => {
  it("splits a named Excel Table without including cells outside its range", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "clients.xlsx");
    const output = path.join(directory, "split");
    await copyWorkbookWithExcelTable(input);

    const tables = await readWorkbookExcelTables(input);
    expect(tables).toHaveLength(1);
    expect(tables[0]).toMatchObject({
      columns: ["Client", "Region", "Amount"],
      excelTableName: "ClientData",
      excelTableRange: "B4:D8",
      source: {
        file: "clients.xlsx",
        firstDataRow: 5,
        sheet: "Clients",
      },
      sourceRows: [5, 6, 7],
    });

    const result = await splitWorkbookByColumn({
      input: input,
      outputDirectory: output,
      column: "Region",
      preserveWorkbook: false,
      table: "clientdata",
    });

    expect(result.metrics).toEqual({
      groups: 2,
      inputFiles: 1,
      inputRows: 3,
      outputFiles: 2,
      outputRows: 3,
      skippedRows: 0,
    });
    expect(
      await readRows(path.join(output, "clients-North.xlsx"), "Clients"),
    ).toEqual([
      { Amount: 10, Client: "A", Region: "North" },
      { Amount: 30, Client: "C", Region: "North" },
    ]);
    expect(
      await readRows(path.join(output, "clients-South.xlsx"), "Clients"),
    ).toEqual([{ Amount: 20, Client: "B", Region: "South" }]);
  });

  it("preserves the workbook by default and changes only the selected Excel Table rows", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "clients.xlsx");
    const output = path.join(directory, "preserved");
    await copyWorkbookWithExcelTable(input);

    const inputArchive = await JSZip.loadAsync(await readFile(input));
    const result = await splitWorkbookByColumn({
      input: input,
      outputDirectory: output,
      column: "Region",
      table: "ClientData",
    });

    expect(result.metrics).toMatchObject({
      groups: 2,
      inputRows: 3,
      outputFiles: 2,
      outputRows: 3,
    });

    const northPath = path.join(output, "clients-North.xlsx");
    const southPath = path.join(output, "clients-South.xlsx");
    const northTables = await readWorkbookExcelTables(northPath);
    const southTables = await readWorkbookExcelTables(southPath);
    expect(northTables).toMatchObject([
      {
        excelTableName: "ClientData",
        excelTableRange: "B4:D7",
        rows: [
          { Amount: 10, Client: "A", Region: "North" },
          { Amount: 30, Client: "C", Region: "North" },
        ],
        sourceRows: [5, 6],
      },
    ]);
    expect(southTables).toMatchObject([
      {
        excelTableName: "ClientData",
        excelTableRange: "B4:D6",
        rows: [{ Amount: 20, Client: "B", Region: "South" }],
        sourceRows: [5],
      },
    ]);

    const northArchive = await JSZip.loadAsync(await readFile(northPath));
    for (const part of [
      "xl/styles.xml",
      "xl/theme/theme1.xml",
      "xl/workbook.xml",
      "xl/worksheets/sheet1.xml",
    ]) {
      expect(await northArchive.file(part)?.async("text")).toBe(
        await inputArchive.file(part)?.async("text"),
      );
    }

    const northWorksheetXml = await northArchive
      .file("xl/worksheets/sheet2.xml")!
      .async("text");
    expect(northWorksheetXml).toContain("Client allocation report");
    expect(northWorksheetXml).toContain("Cells outside ClientData");
    expect(northWorksheetXml).toContain('<x:mergeCell ref="A1:D2" />');
    expect(northWorksheetXml).toContain(
      '<x:c r="D6" s="6" t="n"><x:v>30</x:v></x:c>',
    );
    expect(northWorksheetXml).toContain(
      '<x:c r="D7" s="6" t="n"><x:v>60</x:v></x:c>',
    );
    expect(northWorksheetXml).toContain('<x:c r="D8" s="6" />');

    const northTableXml = await northArchive
      .file("xl/tables/table1.xml")!
      .async("text");
    expect(northTableXml).toContain('ref="B4:D7"');
    expect(northTableXml).toContain('name="TableStyleMedium2"');
  });

  it("reports invalid Excel Table selections and header overrides", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "clients.xlsx");
    await copyWorkbookWithExcelTable(input);

    await expect(
      splitWorkbookByColumn({
        input: input,
        outputDirectory: path.join(directory, "missing"),
        column: "Region",
        table: "MissingTable",
      }),
    ).rejects.toThrowError(
      /Excel Table "MissingTable" was not found or has no data rows/,
    );

    await expect(
      splitWorkbookByColumn({
        input: input,
        outputDirectory: path.join(directory, "invalid-header"),
        column: "Region",
        headerRow: 4,
        table: "ClientData",
      }),
    ).rejects.toThrowError(
      /headerRow option cannot be used with an Excel Table/,
    );

    await expect(
      splitWorkbookByColumn({
        input: input,
        outputDirectory: path.join(directory, "preserve-sheet"),
        column: "Region",
        preserveWorkbook: true,
      }),
    ).rejects.toThrowError(
      /preserveWorkbook option requires a named Excel Table/,
    );
  });

  it("requires an explicit opt-in for Excel Tables on hidden worksheets", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "clients.xlsx");
    await copyWorkbookWithExcelTable(input, true);

    expect(await readWorkbookExcelTables(input)).toEqual([]);
    expect(
      await readWorkbookExcelTables(input, { includeHiddenSheets: true }),
    ).toHaveLength(1);

    await expect(
      splitWorkbookByColumn({
        input: input,
        outputDirectory: path.join(directory, "default"),
        column: "Region",
        table: "ClientData",
      }),
    ).rejects.toThrowError(
      /Excel Table "ClientData" was not found or has no data rows/,
    );

    await expect(
      splitWorkbookByColumn({
        input: input,
        outputDirectory: path.join(directory, "included"),
        column: "Region",
        includeHiddenSheets: true,
        table: "ClientData",
      }),
    ).resolves.toMatchObject({
      metrics: {
        inputRows: 3,
        outputFiles: 2,
      },
    });
  });

  it("writes one workbook per typed value with safe, stable filenames", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "clients.xlsx");
    const output = path.join(directory, "split");
    await createWorkbook(input, [
      [
        "Clients",
        [
          ["Client", "Region", "Amount"],
          ["A", "North", 10],
          ["B", "South", 20],
          ["C", "North", 30],
          ["D", null, 40],
          ["E", "North/West", 50],
          ["F", "North:West", 60],
          ["G", 1, 70],
          ["H", "1", 80],
        ],
      ],
    ]);

    const result = await splitWorkbookByColumn({
      input: input,
      outputDirectory: output,
      column: " region ",
    });

    expect(result.metrics).toEqual({
      groups: 7,
      inputFiles: 1,
      inputRows: 8,
      outputFiles: 7,
      outputRows: 8,
      skippedRows: 0,
    });
    expect(
      result.artifacts.map((artifact) => path.basename(artifact.path)),
    ).toEqual([
      "clients-North.xlsx",
      "clients-South.xlsx",
      "clients-blank.xlsx",
      "clients-North-West.xlsx",
      "clients-North-West-2.xlsx",
      "clients-1.xlsx",
      "clients-1-2.xlsx",
    ]);
    expect(
      await readRows(path.join(output, "clients-North.xlsx"), "Clients"),
    ).toEqual([
      { Amount: 10, Client: "A", Region: "North" },
      { Amount: 30, Client: "C", Region: "North" },
    ]);
    expect(
      await readRows(path.join(output, "clients-blank.xlsx"), "Clients"),
    ).toEqual([{ Amount: 40, Client: "D", Region: null }]);
    expect(
      await readRows(path.join(output, "clients-1.xlsx"), "Clients"),
    ).toEqual([{ Amount: 70, Client: "G", Region: 1 }]);
    expect(
      await readRows(path.join(output, "clients-1-2.xlsx"), "Clients"),
    ).toEqual([{ Amount: 80, Client: "H", Region: "1" }]);
  });

  it("requires an explicit worksheet when more than one table is available", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "regions.xlsx");
    await createWorkbook(input, [
      [
        "Current",
        [
          ["Client", "Region"],
          ["A", "North"],
        ],
      ],
      [
        "Archive",
        [
          ["Client", "Region"],
          ["B", "South"],
        ],
      ],
    ]);

    await expect(
      splitWorkbookByColumn({
        input: input,
        outputDirectory: path.join(directory, "ambiguous"),
        column: "Region",
      }),
    ).rejects.toThrowError(/multiple non-empty worksheets/);

    const result = await splitWorkbookByColumn({
      input: input,
      outputDirectory: path.join(directory, "selected"),
      column: "Region",
      sheet: "archive",
    });
    expect(result.artifacts).toHaveLength(1);
    expect(await readRows(result.artifacts[0]!.path, "Archive")).toEqual([
      { Client: "B", Region: "South" },
    ]);
  });

  it("preflights every output before writing and can skip blank values", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "clients.xlsx");
    const output = path.join(directory, "split");
    await createWorkbook(input, [
      [
        "Clients",
        [
          ["Client", "Region"],
          ["A", "North"],
          ["B", null],
          ["C", "South"],
        ],
      ],
    ]);

    await mkdir(output);
    await createWorkbook(path.join(output, "clients-South.xlsx"), [
      [
        "Existing",
        [
          ["Do", "Not"],
          ["Replace", "Me"],
        ],
      ],
    ]);

    await expect(
      splitWorkbookByColumn({
        input: input,
        outputDirectory: output,
        column: "Region",
        includeBlank: false,
      }),
    ).rejects.toThrowError(/Output already exists/);
    await expect(
      readFile(path.join(output, "clients-North.xlsx")),
    ).rejects.toThrow();

    const result = await splitWorkbookByColumn({
      input: input,
      outputDirectory: output,
      column: "Region",
      includeBlank: false,
      overwrite: true,
    });
    expect(result.metrics.skippedRows).toBe(1);
    expect(result.warnings).toEqual([
      'Skipped 1 row with blank values in "Region".',
    ]);
    expect(
      await readRows(path.join(output, "clients-South.xlsx"), "Clients"),
    ).toEqual([{ Client: "C", Region: "South" }]);
    expect(
      (await readdir(output)).some((filename) =>
        filename.startsWith(".consultchimps-split-"),
      ),
    ).toBe(false);
  });

  it("rejects non-file destinations before creating any output", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "clients.xlsx");
    const output = path.join(directory, "split");
    await createWorkbook(input, [
      [
        "Clients",
        [
          ["Client", "Region"],
          ["A", "North"],
          ["B", "South"],
        ],
      ],
    ]);
    await mkdir(path.join(output, "clients-South.xlsx"), { recursive: true });

    await expect(
      splitWorkbookByColumn({
        input: input,
        outputDirectory: output,
        column: "Region",
        overwrite: true,
      }),
    ).rejects.toThrowError(/exists but is not a file/);
    await expect(
      readFile(path.join(output, "clients-North.xlsx")),
    ).rejects.toThrow();
  });

  it("refuses to relocate position-dependent formulas during a preserved split", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "clients.xlsx");
    const output = path.join(directory, "split");
    await copyWorkbookWithExcelTable(input);

    const archive = await JSZip.loadAsync(await readFile(input));
    const sheetXml = await archive
      .file("xl/worksheets/sheet2.xml")!
      .async("text");
    archive.file(
      "xl/worksheets/sheet2.xml",
      sheetXml.replace(
        '<x:c r="D7" s="6" t="n"><x:v>30</x:v></x:c>',
        '<x:c r="D7" s="6" t="n"><x:f>D5*3</x:f><x:v>30</x:v></x:c>',
      ),
    );
    await writeFile(
      input,
      await archive.generateAsync({
        compression: "DEFLATE",
        type: "nodebuffer",
      }),
    );

    await expect(
      splitWorkbookByColumn({
        input,
        outputDirectory: output,
        column: "Region",
        table: "ClientData",
      }),
    ).rejects.toMatchObject({ code: "XLSX_SPLIT_PRESERVE_FORMULA" });
    await expect(readdir(output)).resolves.toEqual([]);

    const compact = await splitWorkbookByColumn({
      input,
      outputDirectory: output,
      column: "Region",
      preserveWorkbook: false,
      table: "ClientData",
    });
    expect(compact.metrics.outputFiles).toBe(2);
  });

  it("relocates position-independent structured formulas during a preserved split", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "clients.xlsx");
    const output = path.join(directory, "split");
    await copyWorkbookWithExcelTable(input);

    const archive = await JSZip.loadAsync(await readFile(input));
    const sheetXml = await archive
      .file("xl/worksheets/sheet2.xml")!
      .async("text");
    archive.file(
      "xl/worksheets/sheet2.xml",
      sheetXml.replace(
        '<x:c r="D7" s="6" t="n"><x:v>30</x:v></x:c>',
        '<x:c r="D7" s="6" t="n"><x:f>ClientData[[#This Row],[Amount]]</x:f><x:v>30</x:v></x:c>',
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
      input,
      outputDirectory: output,
      column: "Region",
      table: "ClientData",
    });
    expect(result.metrics.outputFiles).toBe(2);

    const northArchive = await JSZip.loadAsync(
      await readFile(path.join(output, "clients-North.xlsx")),
    );
    const northSheetXml = await northArchive
      .file("xl/worksheets/sheet2.xml")!
      .async("text");
    expect(northSheetXml).toContain(
      '<x:c r="D6" s="6" t="n"><x:f>ClientData[[#This Row],[Amount]]</x:f><x:v>30</x:v></x:c>',
    );
  });

  it("plans the split without creating any file and flags collisions", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "clients.xlsx");
    const output = path.join(directory, "split");
    await createWorkbook(input, [
      [
        "Clients",
        [
          ["Client", "Region"],
          ["A", "North"],
          ["B", "South"],
          ["C", null],
        ],
      ],
    ]);

    const plan = await planSplitWorkbookByColumn({
      input,
      outputDirectory: output,
      column: "Region",
    });
    expect(plan.operation).toBe("sheets.split-by-column");
    expect(plan.inputs).toEqual([path.resolve(input)]);
    expect(plan.outputs.map((planned) => path.basename(planned.path))).toEqual([
      "clients-North.xlsx",
      "clients-South.xlsx",
      "clients-blank.xlsx",
    ]);
    expect(plan.outputs.every((planned) => planned.exists === false)).toBe(
      true,
    );
    expect(plan.metrics).toMatchObject({ groups: 3, inputRows: 3 });
    await expect(readdir(output)).rejects.toThrow();

    await mkdir(output, { recursive: true });
    await createWorkbook(path.join(output, "clients-North.xlsx"), [
      ["Existing", [["Header"]]],
    ]);
    const collidingPlan = await planSplitWorkbookByColumn({
      input,
      outputDirectory: output,
      column: "Region",
    });
    expect(
      collidingPlan.outputs.find(
        (planned) => path.basename(planned.path) === "clients-North.xlsx",
      )?.exists,
    ).toBe(true);
    expect(
      collidingPlan.warnings.some((warning) =>
        warning.includes("already exists"),
      ),
    ).toBe(true);
  });
});
