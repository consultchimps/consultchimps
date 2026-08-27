import { copyFile, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { isConsultChimpsError, OPERATION_ABORTED } from "@consultchimps/core";
import type { OperationProgress } from "@consultchimps/core";
import JSZip from "jszip";
import { afterEach, describe, expect, it } from "vitest";
import * as XLSX from "xlsx";

import {
  describeWorkbookBytes,
  readWorkbookExcelTablesBytes,
  readWorkbookNamedRangesBytes,
} from "../src/bytes.js";
import {
  describeWorkbook,
  readWorkbookExcelTables,
  readWorkbookNamedRanges,
} from "../src/index.js";

const temporaryDirectories: string[] = [];
const structuredTableFixture = fileURLToPath(
  new URL("./fixtures/structured-table.xlsx", import.meta.url),
);

async function createTemporaryDirectory(): Promise<string> {
  const directory = await mkdtemp(
    path.join(tmpdir(), "consultchimps-xlsx-describe-"),
  );
  temporaryDirectories.push(directory);
  return directory;
}

interface SheetSpec {
  /** 0 visible, 1 hidden, 2 very hidden — the states Excel itself records. */
  hidden?: 0 | 1 | 2;
  name: string;
  rows: Array<Array<boolean | null | number | string>>;
}

/**
 * A neutral synthetic workbook in the house fixture vocabulary: review-log
 * columns and compass-point regions, never a real organization's data.
 */
function workbookBytes(
  sheets: SheetSpec[],
  names?: Array<{ Name: string; Ref: string }>,
): Uint8Array {
  const workbook = XLSX.utils.book_new();
  for (const sheet of sheets) {
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet(sheet.rows),
      sheet.name,
    );
  }
  workbook.Workbook = {
    Sheets: sheets.map((sheet) => ({
      name: sheet.name,
      Hidden: sheet.hidden ?? 0,
    })),
    ...(names ? { Names: names } : {}),
  };
  return new Uint8Array(
    XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }) as Buffer,
  );
}

async function writeWorkbook(
  filePath: string,
  sheets: SheetSpec[],
  names?: Array<{ Name: string; Ref: string }>,
): Promise<Uint8Array> {
  const bytes = workbookBytes(sheets, names);
  await writeFile(filePath, bytes);
  return bytes;
}

const REVIEW_LOG: SheetSpec = {
  name: "Review Log",
  rows: [
    ["Case_ID", "Region", "Failed Checks"],
    ["R-1", "north", 5],
    ["R-2", "south", 7],
    ["R-3", "north", 9],
  ],
};

afterEach(async () => {
  await Promise.all(
    temporaryDirectories
      .splice(0)
      .map((directory) => rm(directory, { force: true, recursive: true })),
  );
});

describe("describeWorkbook", () => {
  it("describes worksheets, columns, and bounded samples", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "north.xlsx");
    await writeWorkbook(input, [REVIEW_LOG]);

    const { description, result } = await describeWorkbook(input);

    expect(description.source).toBe("north.xlsx");
    expect(description.sheets).toEqual([
      {
        name: "Review Log",
        visibility: "visible",
        rowCount: 4,
        columnCount: 3,
        headerRow: 1,
        dataRowCount: 3,
        columns: [
          {
            header: "Case_ID",
            index: 0,
            sampleValues: ["R-1", "R-2", "R-3"],
          },
          {
            header: "Region",
            index: 1,
            // Distinct values only: "north" appears twice and is sampled once.
            sampleValues: ["north", "south"],
          },
          { header: "Failed Checks", index: 2, sampleValues: [5, 7, 9] },
        ],
      },
    ]);

    // An inspection creates nothing, so the result carries no artifacts.
    expect(result.operation).toBe("sheets.inspect");
    expect(result.artifacts).toEqual([]);
    expect(result.warnings).toEqual([]);
    expect(result.metrics).toEqual({
      dataRows: 3,
      excelTables: 0,
      headerColumns: 3,
      hiddenWorksheets: 0,
      namedRanges: 0,
      worksheets: 1,
    });
  });

  it("describes the file and byte surfaces identically", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "north.xlsx");
    const bytes = await writeWorkbook(
      input,
      [
        REVIEW_LOG,
        {
          name: "Lookup",
          rows: [
            ["Case_ID", "Owner"],
            ["R-4", "Reviewer A"],
          ],
        },
      ],
      [{ Name: "CaseRange", Ref: "Lookup!$A$1:$B$2" }],
    );

    const fromFile = await describeWorkbook(input);
    const fromBytes = await describeWorkbookBytes({
      name: "north.xlsx",
      bytes,
    });

    expect(fromBytes.description).toEqual(fromFile.description);
    expect(fromBytes.result).toEqual(fromFile.result);
    expect(fromFile.description.namedRanges).toEqual([
      { name: "CaseRange", ref: "A1:B2", sheet: "Lookup" },
    ]);
    expect(fromFile.result.metrics.worksheets).toBe(2);
  });

  it("reports Excel Tables with the headers the table part declares", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "clients.xlsx");
    await copyFile(structuredTableFixture, input);

    const { description, result } = await describeWorkbook(input);

    expect(description.excelTables).toEqual([
      {
        name: "ClientData",
        range: "B4:D8",
        sheet: "Clients",
        headers: ["Client", "Region", "Amount"],
      },
    ]);
    expect(result.metrics.excelTables).toBe(1);
    expect(description.sheets.map((sheet) => sheet.name)).toEqual([
      "Cover",
      "Clients",
    ]);
  });

  it("honours a custom header row in the preview", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "north.xlsx");
    await writeWorkbook(input, [
      {
        name: "Review Log",
        rows: [
          ["Quarterly review log", null, null],
          ["Case_ID", "Region", "Failed Checks"],
          ["R-1", "north", 5],
          ["R-2", "south", 7],
        ],
      },
    ]);

    // Without a header row the first row carrying any value wins, which is the
    // title - exactly the confusion an inspection exists to surface.
    const detected = await describeWorkbook(input);
    expect(detected.description.sheets[0]?.headerRow).toBe(1);
    expect(
      detected.description.sheets[0]?.columns.map((column) => column.header),
    ).toEqual(["Quarterly review log", "column_2", "column_3"]);

    const configured = await describeWorkbook(input, { headerRow: 2 });
    const sheet = configured.description.sheets[0];
    expect(sheet?.headerRow).toBe(2);
    expect(sheet?.columns.map((column) => column.header)).toEqual([
      "Case_ID",
      "Region",
      "Failed Checks",
    ]);
    expect(sheet?.dataRowCount).toBe(2);
    // The used range is unchanged by the header choice; only the preview moves.
    expect(sheet?.rowCount).toBe(4);
  });

  it("bounds sample values and rejects an out-of-range request", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "north.xlsx");
    await writeWorkbook(input, [
      {
        name: "Review Log",
        rows: [
          ["Case_ID"],
          ...Array.from({ length: 12 }, (_, index) => [`R-${index + 1}`]),
        ],
      },
    ]);

    const capped = await describeWorkbook(input);
    expect(capped.description.sheets[0]?.columns[0]?.sampleValues).toEqual([
      "R-1",
      "R-2",
      "R-3",
      "R-4",
      "R-5",
    ]);
    expect(capped.description.sheets[0]?.dataRowCount).toBe(12);

    const narrowed = await describeWorkbook(input, { sampleValues: 2 });
    expect(narrowed.description.sheets[0]?.columns[0]?.sampleValues).toEqual([
      "R-1",
      "R-2",
    ]);

    const none = await describeWorkbook(input, { sampleValues: 0 });
    expect(none.description.sheets[0]?.columns[0]?.sampleValues).toEqual([]);
    // Suppressing samples must not change the structure around them.
    expect(none.description.sheets[0]?.dataRowCount).toBe(12);

    for (const sampleValues of [6, -1, 1.5]) {
      await expect(
        describeWorkbook(input, { sampleValues }),
      ).rejects.toMatchObject({ code: "XLSX_INVALID_SAMPLE_LIMIT" });
    }
  });

  it("distinguishes stored values that only look alike", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "north.xlsx");
    await writeWorkbook(input, [
      {
        name: "Review Log",
        rows: [["Failed Checks"], [1], ["1"], [true], ["1"]],
      },
    ]);

    // The number 1 and the text "1" are different stored values, and a mapping
    // review has to see both; the text repeat is still only sampled once.
    expect(
      (await describeWorkbook(input)).description.sheets[0]?.columns[0]
        ?.sampleValues,
    ).toEqual([1, "1", true]);
  });

  it("produces identical descriptions for identical inputs and never writes", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "north.xlsx");
    const bytes = await writeWorkbook(input, [REVIEW_LOG]);
    const originalBytes = Uint8Array.from(bytes);
    const originalFile = await readFile(input);

    const first = await describeWorkbook(input);
    const second = await describeWorkbook(input);
    expect(second).toEqual(first);

    const firstBytes = await describeWorkbookBytes({
      name: "north.xlsx",
      bytes,
    });
    const secondBytes = await describeWorkbookBytes({
      name: "north.xlsx",
      bytes,
    });
    expect(secondBytes).toEqual(firstBytes);

    // Describing reads only: the input bytes and the file on disk are untouched.
    expect(bytes.length).toBe(originalBytes.length);
    expect(Buffer.compare(Buffer.from(bytes), Buffer.from(originalBytes))).toBe(
      0,
    );
    expect(Buffer.compare(await readFile(input), originalFile)).toBe(0);
  });

  it("reports deterministic progress and honours cancellation", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "north.xlsx");
    await writeWorkbook(input, [
      REVIEW_LOG,
      { name: "Lookup", rows: [["Case_ID"], ["R-4"]] },
    ]);

    const events: OperationProgress[] = [];
    await describeWorkbook(input, {
      onProgress: (progress) => events.push(progress),
    });
    expect(events.map((event) => [event.stage, event.completed])).toEqual([
      ["describing-worksheets", 1],
      ["describing-worksheets", 2],
      ["describing-structures", 1],
    ]);
    expect(events.every((event) => event.operation === "sheets.inspect")).toBe(
      true,
    );

    const controller = new AbortController();
    controller.abort();
    let thrown: unknown;
    try {
      await describeWorkbook(input, { signal: controller.signal });
    } catch (error) {
      thrown = error;
    }
    expect(isConsultChimpsError(thrown)).toBe(true);
    expect((thrown as { code: string }).code).toBe(OPERATION_ABORTED);
  });
});

describe("describeWorkbook hidden worksheets", () => {
  const sheets: SheetSpec[] = [
    REVIEW_LOG,
    { hidden: 1, name: "Summary", rows: [["Category"], ["Complete"]] },
    { hidden: 2, name: "Vault", rows: [["Case_ID"], ["R-9"]] },
  ];

  it("excludes hidden worksheets by default and says so", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "north.xlsx");
    await writeWorkbook(input, sheets);

    const { description, result } = await describeWorkbook(input);

    expect(description.sheets.map((sheet) => sheet.name)).toEqual([
      "Review Log",
    ]);
    expect(result.metrics.worksheets).toBe(1);
    expect(result.metrics.hiddenWorksheets).toBe(0);
    expect(result.warnings).toEqual([
      "2 worksheets are hidden and were not described. Include hidden worksheets to describe them.",
    ]);
  });

  it("reports every visibility state when hidden worksheets are included", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "north.xlsx");
    await writeWorkbook(input, sheets);

    const { description, result } = await describeWorkbook(input, {
      includeHiddenSheets: true,
    });

    expect(
      description.sheets.map((sheet) => [sheet.name, sheet.visibility]),
    ).toEqual([
      ["Review Log", "visible"],
      ["Summary", "hidden"],
      ["Vault", "very-hidden"],
    ]);
    expect(result.metrics.hiddenWorksheets).toBe(2);
    expect(result.warnings).toEqual([]);
  });

  it("selects named worksheets case-insensitively", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "north.xlsx");
    await writeWorkbook(input, sheets);

    const { description } = await describeWorkbook(input, {
      includeHiddenSheets: true,
      sheets: ["vault"],
    });
    expect(description.sheets.map((sheet) => sheet.name)).toEqual(["Vault"]);
  });
});

describe("describeWorkbook expected failures", () => {
  it("refuses a workbook that declares no worksheets", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "empty.xlsx");
    const archive = await JSZip.loadAsync(
      Buffer.from(workbookBytes([REVIEW_LOG])),
    );
    const workbookPart = archive.file("xl/workbook.xml");
    const workbookXml = await workbookPart!.async("text");
    const strippedXml = workbookXml.replace(
      /<sheets>.*?<\/sheets>/su,
      "<sheets/>",
    );
    expect(strippedXml).not.toBe(workbookXml);
    archive.file("xl/workbook.xml", strippedXml);
    const bytes = await archive.generateAsync({ type: "nodebuffer" });
    await writeFile(input, bytes);

    await expect(describeWorkbook(input)).rejects.toMatchObject({
      code: "XLSX_NO_SHEETS",
    });
    await expect(
      describeWorkbookBytes({ name: "empty.xlsx", bytes }),
    ).rejects.toMatchObject({ code: "XLSX_NO_SHEETS" });
  });

  it("refuses a worksheet selection the workbook does not have", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "north.xlsx");
    const bytes = await writeWorkbook(input, [REVIEW_LOG]);

    await expect(
      describeWorkbook(input, { sheets: ["Review Log", "Missing"] }),
    ).rejects.toMatchObject({
      code: "XLSX_WORKSHEET_NOT_FOUND",
      details: { missingWorksheets: ["Missing"] },
    });
    await expect(
      describeWorkbookBytes(
        { name: "north.xlsx", bytes },
        {
          sheets: ["Missing"],
        },
      ),
    ).rejects.toMatchObject({ code: "XLSX_WORKSHEET_NOT_FOUND" });
  });

  it("reports an unreadable workbook with the shared read error", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "missing.xlsx");

    await expect(describeWorkbook(input)).rejects.toMatchObject({
      code: "XLSX_READ_FAILED",
    });
    await expect(
      describeWorkbookBytes({
        name: "broken.xlsx",
        bytes: new Uint8Array([1, 2, 3, 4]),
      }),
    ).rejects.toMatchObject({ code: "XLSX_READ_FAILED" });
  });

  it("describes a worksheet with no values instead of refusing", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "north.xlsx");
    await writeWorkbook(input, [{ name: "Blank", rows: [[]] }]);

    const { description, result } = await describeWorkbook(input);
    expect(description.sheets).toEqual([
      {
        name: "Blank",
        visibility: "visible",
        rowCount: 0,
        columnCount: 0,
        headerRow: undefined,
        columns: [],
        dataRowCount: 0,
      },
    ]);
    expect(result.warnings).toEqual([
      'No header row was found in "Blank". An operation that matches columns by header would find nothing to match in it.',
    ]);
  });
});

describe("byte-surface readers match the file surface", () => {
  it("reads the same Excel Tables from bytes as from a path", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "clients.xlsx");
    await copyFile(structuredTableFixture, input);
    const bytes = new Uint8Array(await readFile(input));

    const fromFile = await readWorkbookExcelTables(input);
    const fromBytes = await readWorkbookExcelTablesBytes({
      name: "clients.xlsx",
      bytes,
    });

    expect(fromBytes).toEqual(fromFile);
    expect(fromBytes[0]).toMatchObject({
      columns: ["Client", "Region", "Amount"],
      excelTableName: "ClientData",
      excelTableRange: "B4:D8",
    });

    // The selection options travel with the reader.
    expect(
      await readWorkbookExcelTablesBytes(
        { name: "clients.xlsx", bytes },
        { tables: ["missing"] },
      ),
    ).toEqual([]);
  });

  it("reads the same named ranges from bytes as from a path", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "north.xlsx");
    const bytes = await writeWorkbook(
      input,
      [REVIEW_LOG],
      [
        { Name: "CaseRange", Ref: "'Review Log'!$A$1:$C$4" },
        { Name: "_xlnm.Print_Area", Ref: "'Review Log'!$A$1:$C$4" },
      ],
    );

    const fromFile = await readWorkbookNamedRanges(input);
    const fromBytes = await readWorkbookNamedRangesBytes({
      name: "north.xlsx",
      bytes,
    });

    expect(fromBytes).toEqual(fromFile);
    expect(fromBytes).toHaveLength(1);
    expect(fromBytes[0]).toMatchObject({
      columns: ["Case_ID", "Region", "Failed Checks"],
      rangeName: "CaseRange",
      rangeRef: "A1:C4",
    });

    expect(
      await readWorkbookNamedRangesBytes(
        { name: "north.xlsx", bytes },
        { names: ["missing"] },
      ),
    ).toEqual([]);
  });
});
