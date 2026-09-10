import { copyFile, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { isConsultChimpsError, OPERATION_ABORTED } from "@consultchimps/core";
import type { OperationProgress } from "@consultchimps/core";
import JSZip from "jszip";
import { afterEach, describe, expect, it } from "vitest";
import * as XLSX from "xlsx";

import { XLSX_ERRORS } from "../src/errors.js";
import { forEachZone, ZONES } from "./zones.js";
import {
  describeWorkbookBytes,
  readWorkbookExcelTablesBytes,
  readWorkbookNamedRangesBytes,
  readWorkbookTablesBytes,
  readWorkbookWorksheetsBytes,
} from "../src/bytes.js";
import {
  describeWorkbook,
  readWorkbookExcelTables,
  readWorkbookNamedRanges,
  readWorkbookTables,
  readWorkbookWorksheets,
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
  /** 0 visible, 1 hidden, 2 very hidden: the states Excel itself records. */
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

/**
 * A small worksheet followed by a densely populated one: every cell of every
 * row is present, which is the shape that made per-coordinate cell lookup
 * quadratic. The first sheet exists so a test can act once it is described.
 */
function denseWorkbookBytes(rows: number): Uint8Array {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([["Case_ID"], ["R-0"]]),
    "Review Log",
  );
  const dense: Array<Array<number | string>> = [
    ["Case_ID", "Region", "Failed Checks", "Total Checks", "Owner"],
  ];
  for (let row = 0; row < rows; row += 1) {
    dense.push([
      `R-${row + 1}`,
      row % 2 === 0 ? "north" : "south",
      row % 7,
      100,
      `Reviewer ${row % 5}`,
    ]);
  }
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet(dense),
    "Dense",
  );
  return new Uint8Array(
    XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }) as Buffer,
  );
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

  it("collects a cancellation posted while the scan is running", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "north.xlsx");
    // Several worksheets, each long enough to cross the in-scan yield, so the
    // abort below has to be collected mid-operation rather than before it.
    await writeWorkbook(
      input,
      Array.from({ length: 4 }, (_, sheet) => ({
        name: `Region ${sheet + 1}`,
        rows: [
          ["Case_ID", "Region", "Failed Checks"],
          ...Array.from({ length: 1500 }, (_, row) => [
            `R-${row + 1}`,
            row % 2 === 0 ? "north" : "south",
            row % 11,
          ]),
        ],
      })),
    );

    // The abort is posted as a macrotask from inside the first progress event,
    // so it is queued only once the workbook is loaded and a worksheet has
    // actually been described. An inline `controller.abort()` before the call
    // would only prove the entry check works.
    //
    // What makes this a regression test is the count below rather than the
    // throw: a scan that never yields cannot dequeue the abort until every
    // worksheet is finished, so it would still throw - after reporting all four
    // worksheets. Collecting the cancellation early is the behaviour at stake.
    const controller = new AbortController();
    const described: string[] = [];
    let thrown: unknown;
    try {
      await describeWorkbook(input, {
        signal: controller.signal,
        onProgress: (progress) => {
          if (progress.stage === "describing-worksheets") {
            described.push(progress.detail ?? "");
            if (described.length === 1) {
              setTimeout(() => controller.abort(), 0);
            }
          }
        },
      });
    } catch (error) {
      thrown = error;
    }

    expect(isConsultChimpsError(thrown)).toBe(true);
    expect((thrown as { code: string }).code).toBe(OPERATION_ABORTED);
    expect(described.length).toBeGreaterThan(0);
    // Stopped well before the fourth worksheet, so the cancellation was
    // collected mid-operation rather than after all the work was already done.
    expect(described.length).toBeLessThan(3);
  });

  it("refuses an already-aborted signal before touching the filesystem", async () => {
    const directory = await createTemporaryDirectory();
    const missing = path.join(directory, "does-not-exist.xlsx");
    const controller = new AbortController();
    controller.abort();

    // The file does not exist, so reading it would fail with XLSX_READ_FAILED.
    // A caller who already cancelled must be told they cancelled: the abort is
    // checked before any filesystem work, exactly as the byte twin does.
    await expect(
      describeWorkbook(missing, { signal: controller.signal }),
    ).rejects.toMatchObject({ code: OPERATION_ABORTED });

    // And the same for a workbook that does exist, so the check is not merely
    // shadowing a read error.
    const present = path.join(directory, "north.xlsx");
    await writeWorkbook(present, [REVIEW_LOG]);
    await expect(
      describeWorkbook(present, { signal: controller.signal }),
    ).rejects.toMatchObject({ code: OPERATION_ABORTED });
  });

  it("describes a formatting-only template as empty and stays cancellable", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "template.xlsx");

    // A styled template: 30,000 stored `<row>` elements carrying no cells at
    // all. Nothing here is content, so the sheet must describe as empty, and
    // materializing those rows is the synchronous burst the scan yields around.
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet([["Case_ID"], ["R-1"]]),
      "Review Log",
    );
    const template: XLSX.WorkSheet = { "!ref": "A1:B30000" };
    for (let row = 1; row <= 30000; row += 1) {
      template[`A${row}`] = { t: "z", s: 1 } as XLSX.CellObject;
      template[`B${row}`] = { t: "z", s: 1 } as XLSX.CellObject;
    }
    XLSX.utils.book_append_sheet(workbook, template, "Template");
    await writeFile(
      input,
      XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }),
    );

    const { description, result } = await describeWorkbook(input);
    const templateSheet = description.sheets.find(
      (sheet) => sheet.name === "Template",
    );
    expect(templateSheet?.headerRow).toBeUndefined();
    expect(templateSheet?.dataRowCount).toBe(0);
    expect(result.warnings).toEqual([
      'No header row was found in "Template". An operation that matches columns by header would find nothing to match in it.',
    ]);

    // A cancellation queued while the first worksheet is reported is collected
    // before the template is described, rather than after its rows are built.
    const controller = new AbortController();
    const described: string[] = [];
    let thrown: unknown;
    try {
      await describeWorkbook(input, {
        signal: controller.signal,
        onProgress: (progress) => {
          if (progress.stage === "describing-worksheets") {
            described.push(progress.detail ?? "");
            if (described.length === 1) {
              setTimeout(() => controller.abort(), 0);
            }
          }
        },
      });
    } catch (error) {
      thrown = error;
    }
    expect(isConsultChimpsError(thrown)).toBe(true);
    expect((thrown as { code: string }).code).toBe(OPERATION_ABORTED);
    expect(described).toEqual(["Review Log"]);
  }, 60000);

  it("collects a cancellation while scanning a densely populated sheet", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "dense.xlsx");
    await writeFile(input, denseWorkbookBytes(12000));

    // The abort is queued once the first worksheet is described, on a delay
    // long enough to clear the between-worksheet yield but well short of the
    // dense sheet's scan. It can therefore only be collected from inside that
    // scan: without a yield there, the scan runs to the end and "Dense" is
    // reported before the cancellation is ever seen.
    const controller = new AbortController();
    const described: string[] = [];
    let thrown: unknown;
    try {
      await describeWorkbook(input, {
        signal: controller.signal,
        onProgress: (progress) => {
          if (progress.stage === "describing-worksheets") {
            described.push(progress.detail ?? "");
            if (described.length === 1) {
              setTimeout(() => controller.abort(), 100);
            }
          }
        },
      });
    } catch (error) {
      thrown = error;
    }

    expect(isConsultChimpsError(thrown)).toBe(true);
    expect((thrown as { code: string }).code).toBe(OPERATION_ABORTED);
    // Stopped inside the dense sheet's scan, so it was never reported.
    expect(described).toEqual(["Review Log"]);
  }, 60000);

  it("describes a densely populated sheet in time proportional to its contents", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "dense.xlsx");
    await writeFile(input, denseWorkbookBytes(12000));

    const started = Date.now();
    const { description, result } = await describeWorkbook(input);
    const elapsed = Date.now() - started;

    // The dense sheet's 12,000 rows plus the small sheet's single row.
    expect(result.metrics.dataRows).toBe(12001);
    expect(description.sheets[1]?.dataRowCount).toBe(12000);

    // A guard against the cost returning to quadratic, not a benchmark. Cell
    // lookup used to rebuild and rescan the row list per coordinate, so this
    // sheet took minutes; it now runs in a couple of seconds, and the budget
    // is loose enough to absorb a slow CI machine while still failing long
    // before a quadratic scan could finish.
    expect(elapsed).toBeLessThan(30000);
  }, 90000);
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

describe("describeWorkbook stored values", () => {
  it("samples a date-formatted cell as the number the workbook stores", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "north.xlsx");

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet([
      ["Case_ID", "Started"],
      ["R-1", null],
      ["R-2", null],
    ]);
    // A numeric cell carrying a date number format. Excel stores the serial;
    // the date is a presentation choice made by the style.
    worksheet.B2 = { t: "n", v: 45000, z: "yyyy-mm-dd" };
    worksheet.B3 = { t: "n", v: 45001, z: "yyyy-mm-dd" };
    XLSX.utils.book_append_sheet(workbook, worksheet, "Review Log");
    await writeFile(
      input,
      XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }),
    );

    const { description } = await describeWorkbook(input);
    const started = description.sheets[0]?.columns.find(
      (column) => column.header === "Started",
    );

    // The stored value is the serial, so that is what the sample reports. An
    // ISO string here would be an inferred type the ADR excludes, and would
    // show a mapping review a value the cell does not contain.
    expect(started?.sampleValues).toEqual([45000, 45001]);
    for (const value of started?.sampleValues ?? []) {
      expect(typeof value).toBe("number");
    }
  });
});

describe("describeWorkbook row occupancy", () => {
  /**
   * A worksheet whose data rows carry formulas with no cached `<v>` value:
   * what a workbook saved by a tool that does not calculate looks like.
   */
  async function writeUncachedFormulaWorkbook(filePath: string): Promise<void> {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet([
      ["Case_ID", "Doubled"],
      [null, null],
      [null, null],
    ]);
    // Formula cells with no cached result: the cell exists, the row is not
    // blank, but there is no stored value to report.
    worksheet.A2 = { t: "n", f: "1+0" };
    worksheet.B2 = { t: "n", f: "A2*2" };
    worksheet.A3 = { t: "n", f: "2+0" };
    worksheet.B3 = { t: "n", f: "A3*2" };
    XLSX.utils.book_append_sheet(workbook, worksheet, "Review Log");
    await writeFile(
      filePath,
      XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }),
    );
  }

  it("counts a row of uncached formulas as a data row", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "north.xlsx");
    await writeUncachedFormulaWorkbook(input);

    const { description, result } = await describeWorkbook(input);
    const sheet = description.sheets[0];

    // The rows are not blank: the worksheet holds those cells. Reporting zero
    // data rows would tell a reader the sheet is emptier than it is.
    expect(sheet?.dataRowCount).toBe(2);
    expect(result.metrics.dataRows).toBe(2);
    expect(sheet?.headerRow).toBe(1);

    // Occupancy and samples answer different questions: an uncached formula
    // occupies its row but contributes no sample, because there is no stored
    // value and the inspection never computes one.
    for (const column of sheet?.columns ?? []) {
      expect(column.sampleValues).toEqual([]);
    }
  });

  it("describes a sheet of only uncached formulas as populated", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "formulas.xlsx");
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet([[null], [null]]);
    worksheet.A1 = { t: "n", f: "1+1" };
    worksheet.A2 = { t: "n", f: "2+2" };
    XLSX.utils.book_append_sheet(workbook, worksheet, "Formulas");
    await writeFile(
      input,
      XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }),
    );

    const { description, result } = await describeWorkbook(input);

    // The sheet is not empty, so it gets a header row rather than the
    // "nothing here" description reserved for a genuinely blank worksheet.
    expect(description.sheets[0]?.headerRow).toBe(1);
    expect(result.warnings).toEqual([]);
  });

  it("still treats a formatting-only cell as no content", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "styled.xlsx");
    const workbook = XLSX.utils.book_new();
    const worksheet: XLSX.WorkSheet = { "!ref": "A1:A2" };
    // A cell carrying a style and nothing else is formatting, not content.
    worksheet.A1 = { t: "z", s: 1 } as XLSX.CellObject;
    XLSX.utils.book_append_sheet(workbook, worksheet, "Styled");
    await writeFile(
      input,
      XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }),
    );

    const { description } = await describeWorkbook(input);
    expect(description.sheets[0]?.headerRow).toBeUndefined();
    expect(description.sheets[0]?.dataRowCount).toBe(0);
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

  it("refuses an invalid header row whatever the workbook contains", async () => {
    const directory = await createTemporaryDirectory();
    const populated = path.join(directory, "north.xlsx");
    const blank = path.join(directory, "blank.xlsx");
    await writeWorkbook(populated, [REVIEW_LOG]);
    await writeWorkbook(blank, [{ name: "Blank", rows: [[]] }]);
    const hiddenOnly = path.join(directory, "hidden.xlsx");
    await writeWorkbook(hiddenOnly, [
      { hidden: 1, name: "Summary", rows: [["Category"], ["Complete"]] },
    ]);

    // The option is refused identically whether a worksheet has content, has
    // none, or was filtered out as hidden - the workbook never decides whether
    // an option is valid.
    for (const workbook of [populated, blank, hiddenOnly]) {
      for (const headerRow of [0, -1, 1.5]) {
        await expect(
          describeWorkbook(workbook, { headerRow }),
        ).rejects.toMatchObject({ code: "XLSX_INVALID_HEADER_ROW" });
      }
    }
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

describe("describeWorkbook named ranges", () => {
  it("matches a sheet whose name the workbook part has to escape", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "north.xlsx");
    // "&" is stored as &amp; in the defined name's reference, while the sheet
    // name reaches the description decoded. Comparing the two raw dropped the
    // range silently and undercounted the metric.
    const sheetName = "Review & Log";
    await writeWorkbook(
      input,
      [{ name: sheetName, rows: REVIEW_LOG.rows }],
      [{ Name: "CaseRange", Ref: `'${sheetName}'!$A$1:$C$4` }],
    );

    const { description, result } = await describeWorkbook(input);

    expect(description.sheets.map((sheet) => sheet.name)).toEqual([sheetName]);
    expect(description.namedRanges).toEqual([
      { name: "CaseRange", ref: "A1:C4", sheet: sheetName },
    ]);
    expect(result.metrics.namedRanges).toBe(1);
  });
});

/**
 * A worksheet part whose totals column holds formulas nothing has calculated:
 * each cell carries `<f>` and no `<v>`, which is what a generator writes when
 * it has no calculation engine. It is assembled by hand because a spreadsheet
 * engine will not produce it: SheetJS drops a numeric formula cell with no
 * cached value while parsing, which is the whole reason this condition has to
 * be read from the package's own model.
 */
async function uncalculatedWorkbookBytes(rows: string): Promise<Uint8Array> {
  const archive = new JSZip();
  archive.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
  );
  archive.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  );
  archive.file(
    "xl/workbook.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Review Log" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  );
  archive.file(
    "xl/_rels/workbook.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
  );
  archive.file(
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows}</sheetData></worksheet>`,
  );
  return new Uint8Array(await archive.generateAsync({ type: "nodebuffer" }));
}

function textCell(address: string, text: string): string {
  return `<c r="${address}" t="inlineStr"><is><t>${text}</t></is></c>`;
}

describe("dates a workbook stores as numbers", () => {
  /**
   * A worksheet whose columns are date-formatted serials. `date1904` writes the
   * workbook in Excel's other date system, where the same calendar day is 1462
   * serials lower, so the two workbooks below hold different numbers and mean
   * the same days.
   */
  function datedWorkbook(date1904 = false): Uint8Array {
    const shift = date1904 ? 1462 : 0;
    const sheet: XLSX.WorkSheet = {
      "!ref": "A1:E2",
      A1: { t: "s", v: "Customer" },
      B1: { t: "s", v: "Opened" },
      C1: { t: "s", v: "Stamped" },
      D1: { t: "s", v: "Closed" },
      E1: { t: "s", v: "Reference" },
      A2: { t: "s", v: "Acme" },
      // 1 January 2024, and the same day at 18:00.
      B2: { t: "n", v: 45292 - shift, z: "yyyy-mm-dd" },
      C2: { t: "n", v: 45292.75 - shift, z: "yyyy-mm-dd hh:mm:ss" },
      // The last fraction of that day, which rounds into the next one. The
      // carry is arithmetic, so it has to land on the same day in every zone
      // just as the plain serials do.
      D2: {
        t: "n",
        v: 45292 - shift + 86_399_999.6 / 86_400_000,
        z: "yyyy-mm-dd hh:mm:ss",
      },
      // A plain number wearing no date format stays a number, whatever it
      // would decode to if it were read as a serial.
      E2: { t: "n", v: 45292 - shift },
    };
    return new Uint8Array(
      XLSX.write(
        {
          SheetNames: ["Data"],
          Sheets: { Data: sheet },
          Workbook: { WBProps: { date1904 } },
        },
        { bookType: "xlsx", type: "array" },
      ) as ArrayBuffer,
    );
  }

  async function firstRowPerZone(
    bytes: Uint8Array,
  ): Promise<Array<Record<string, unknown> | undefined>> {
    // One zone at a time: the zone is one process-wide value, so two readings
    // in flight together would each restore the other's zone.
    return await forEachZone(async () => {
      const [table] = await readWorkbookTablesBytes({
        name: "dates.xlsx",
        bytes,
      });
      return table?.rows[0];
    });
  }

  it("reads a date-formatted serial the same way in every time zone", async () => {
    // A serial names a calendar moment and carries no zone, so the text is a
    // function of the serial and the workbook's date system and of nothing
    // else. The reader decodes the serial into calendar components directly,
    // with no `Date` in the middle: a `Date` has a local face as well as a UTC
    // one, and text built from the local face made the same workbook read as a
    // different calendar day in every zone.
    expect(await firstRowPerZone(datedWorkbook())).toEqual(
      ZONES.map(() => ({
        Customer: "Acme",
        Opened: "2024-01-01T00:00:00.000Z",
        Stamped: "2024-01-01T18:00:00.000Z",
        Closed: "2024-01-02T00:00:00.000Z",
        Reference: 45292,
      })),
    );
  });

  it("carries a rounded fraction into the next day", async () => {
    // 23:59:59.9996 on 1 January is midnight on the 2nd once it is written to
    // the millisecond. Rounding the fraction after the hour, minute and second
    // were already decided had nowhere to carry, so the value read as the
    // second before, on the day before.
    const almost = 45292 + 86_399_999.6 / 86_400_000;
    const sheet: XLSX.WorkSheet = {
      "!ref": "A1:B4",
      A1: { t: "s", v: "Case" },
      B1: { t: "s", v: "Stamped" },
      A2: { t: "s", v: "R-1" },
      B2: { t: "n", v: almost, z: "yyyy-mm-dd hh:mm:ss" },
      A3: { t: "s", v: "R-2" },
      // Midnight itself, which has nothing to carry.
      B3: { t: "n", v: 45293, z: "yyyy-mm-dd hh:mm:ss" },
      A4: { t: "s", v: "R-3" },
      // Half a millisecond short of a second, which rounds within the day.
      B4: {
        t: "n",
        v: 45292 + 86_399_998.6 / 86_400_000,
        z: "yyyy-mm-dd hh:mm:ss",
      },
    };
    const [table] = await readWorkbookTablesBytes({
      name: "carry.xlsx",
      bytes: new Uint8Array(
        XLSX.write(
          { SheetNames: ["Data"], Sheets: { Data: sheet } },
          { bookType: "xlsx", type: "array" },
        ) as ArrayBuffer,
      ),
    });

    expect(table?.rows).toEqual([
      { Case: "R-1", Stamped: "2024-01-02T00:00:00.000Z" },
      { Case: "R-2", Stamped: "2024-01-02T00:00:00.000Z" },
      { Case: "R-3", Stamped: "2024-01-01T23:59:59.999Z" },
    ]);
  });

  it("carries the same way in the 1904 date system", async () => {
    const sheet: XLSX.WorkSheet = {
      "!ref": "A1:B2",
      A1: { t: "s", v: "Case" },
      B1: { t: "s", v: "Stamped" },
      A2: { t: "s", v: "R-1" },
      B2: {
        t: "n",
        v: 45292 - 1462 + 86_399_999.6 / 86_400_000,
        z: "yyyy-mm-dd hh:mm:ss",
      },
    };
    const [table] = await readWorkbookTablesBytes({
      name: "carry-1904.xlsx",
      bytes: new Uint8Array(
        XLSX.write(
          {
            SheetNames: ["Data"],
            Sheets: { Data: sheet },
            Workbook: { WBProps: { date1904: true } },
          },
          { bookType: "xlsx", type: "array" },
        ) as ArrayBuffer,
      ),
    });

    expect(table?.rows).toEqual([
      { Case: "R-1", Stamped: "2024-01-02T00:00:00.000Z" },
    ]);
  });

  it("keeps a serial that names no calendar day as the number it is", async () => {
    // Serial 0 decodes to day zero of January 1900, which is not a day. There
    // is no date to write, so the number travels and nothing is invented.
    // Serial 60 is a day: the 1900 system deliberately reproduces a
    // spreadsheet-era bug in which 29 February 1900 exists, and that is what
    // Excel shows for it, so that is what the reader writes.
    const sheet: XLSX.WorkSheet = {
      "!ref": "A1:B3",
      A1: { t: "s", v: "Case" },
      B1: { t: "s", v: "Opened" },
      A2: { t: "s", v: "R-1" },
      B2: { t: "n", v: 0, z: "yyyy-mm-dd" },
      A3: { t: "s", v: "R-2" },
      B3: { t: "n", v: 60, z: "yyyy-mm-dd" },
    };
    const [table] = await readWorkbookTablesBytes({
      name: "edges.xlsx",
      bytes: new Uint8Array(
        XLSX.write(
          { SheetNames: ["Data"], Sheets: { Data: sheet } },
          { bookType: "xlsx", type: "array" },
        ) as ArrayBuffer,
      ),
    });

    expect(table?.rows).toEqual([
      { Case: "R-1", Opened: 0 },
      { Case: "R-2", Opened: "1900-02-29T00:00:00.000Z" },
    ]);
  });

  it("reads the 1904 date system by the workbook's own count", async () => {
    // The same days, written in the system that starts 1462 days later. Which
    // day a serial names belongs to the workbook, so the reader asks it.
    expect(await firstRowPerZone(datedWorkbook(true))).toEqual(
      ZONES.map(() => ({
        Customer: "Acme",
        Opened: "2024-01-01T00:00:00.000Z",
        Stamped: "2024-01-01T18:00:00.000Z",
        Closed: "2024-01-02T00:00:00.000Z",
        Reference: 45292 - 1462,
      })),
    );
  });
});

describe("worksheets whose formulas were never calculated", () => {
  const input = async (rows: string) => ({
    name: "north.xlsx",
    bytes: await uncalculatedWorkbookBytes(rows),
  });

  it("counts the cells holding a formula with no calculated value", async () => {
    const [report] = await readWorkbookWorksheetsBytes(
      await input(
        `<row r="1">${textCell("A1", "Case_ID")}${textCell("B1", "Failed Checks")}</row>` +
          `<row r="2">${textCell("A2", "R-1")}<c r="B2"><f>1+1</f></c></row>` +
          `<row r="3">${textCell("A3", "R-2")}<c r="B3"><f>2+2</f></c></row>`,
      ),
    );

    expect(report?.sheet).toBe("Review Log");
    expect(report?.uncachedFormulaCells).toBe(2);
    // The table still reads them as empty, which is all the file says, and is
    // exactly why the count has to travel beside it.
    expect(report?.table?.rows).toEqual([
      { Case_ID: "R-1", "Failed Checks": null },
      { Case_ID: "R-2", "Failed Checks": null },
    ]);
  });

  it("counts a header cell whose formula was never calculated", async () => {
    // The header reads as blank, so the reader invents a name for the column
    // and the worksheet imports under a different schema.
    const [report] = await readWorkbookWorksheetsBytes(
      await input(
        `<row r="1">${textCell("A1", "Case_ID")}<c r="B1"><f>CONCATENATE("Failed"," Checks")</f></c></row>` +
          `<row r="2">${textCell("A2", "R-1")}<c r="B2"><v>5</v></c></row>`,
      ),
    );

    expect(report?.uncachedFormulaCells).toBe(1);
    expect(report?.table?.columns).toEqual(["Case_ID", "column_2"]);
  });

  it("leaves a formula above the header alone, and reads from the header", async () => {
    // The used range opens with a title cell nothing has calculated. The reader
    // sees no value there and keys on the real header below it, so that formula
    // is outside the table that would be imported. Counting it would refuse the
    // worksheet over a cell the import never reads.
    const [report] = await readWorkbookWorksheetsBytes(
      await input(
        `<row r="1"><c r="A1"><f>CONCATENATE("Q","1")</f></c></row>` +
          `<row r="2">${textCell("A2", "Case_ID")}${textCell("B2", "Region")}</row>` +
          `<row r="3">${textCell("A3", "R-1")}${textCell("B3", "north")}</row>`,
      ),
    );

    expect(report?.uncachedFormulaCells).toBe(0);
    // The region reported is the one the rows were read from, header included.
    expect(report?.region?.headerRow).toBe(2);
    expect(report?.table?.source?.firstDataRow).toBe(3);
    expect(report?.table?.rows).toEqual([{ Case_ID: "R-1", Region: "north" }]);
  });

  it("counts the same formula once it sits inside the region", async () => {
    // The identical cell, now under the header rather than above it, is part of
    // what would be imported and is counted.
    const [report] = await readWorkbookWorksheetsBytes(
      await input(
        `<row r="1">${textCell("A1", "Case_ID")}${textCell("B1", "Region")}</row>` +
          `<row r="2">${textCell("A2", "R-1")}<c r="B2"><f>CONCATENATE("Q","1")</f></c></row>`,
      ),
    );

    expect(report?.region?.headerRow).toBe(1);
    expect(report?.uncachedFormulaCells).toBe(1);
  });

  it("counts an uncalculated header in the last column too", async () => {
    const bytes = (
      await input(
        `<row r="1">${textCell("A1", "Case_ID")}<c r="B1"><f>1+1</f></c></row>` +
          `<row r="2">${textCell("A2", "R-1")}<c r="B2"><v>5</v></c></row>`,
      )
    ).bytes;

    const [report] = await readWorkbookWorksheetsBytes({
      name: "north.xlsx",
      bytes,
    });
    expect(report?.uncachedFormulaCells).toBe(1);
    const [table] = await readWorkbookTablesBytes({
      name: "north.xlsx",
      bytes,
    });
    expect(table?.columns).toEqual(["Case_ID", "column_2"]);
  });

  it("treats an empty cached result as calculated, because it is", async () => {
    // A string formula that evaluates to nothing is written with an empty
    // cached value. Reading that as "never calculated" would block the
    // worksheet forever: recalculating in Excel produces the same empty result.
    const [report] = await readWorkbookWorksheetsBytes(
      await input(
        `<row r="1">${textCell("A1", "Case_ID")}${textCell("B1", "Note")}</row>` +
          `<row r="2">${textCell("A2", "R-1")}<c r="B2" t="str"><f>IF(1=2,"x","")</f><v></v></c></row>` +
          `<row r="3">${textCell("A3", "R-2")}<c r="B3"><f>0+0</f><v>0</v></c></row>` +
          `<row r="4">${textCell("A4", "R-3")}<c r="B4"><f>1+1</f><v>  </v></c></row>`,
      ),
    );

    // An empty string, a zero, and whitespace are all calculated results.
    expect(report?.uncachedFormulaCells).toBe(0);
  });

  it("reports the condition even where the reader can build no table", async () => {
    // Every value in the body is a formula with no result, so the worksheet
    // reads as empty and yields no table at all. With no region to scope to,
    // the whole worksheet is counted, which is what explains why it looks
    // empty.
    const bytes = (
      await input(
        `<row r="1">${textCell("A1", "Total")}</row>` +
          `<row r="2"><c r="A2"><f>1+1</f></c></row>`,
      )
    ).bytes;

    const [report] = await readWorkbookWorksheetsBytes({
      name: "summary.xlsx",
      bytes,
    });
    expect(report?.table).toBeUndefined();
    expect(report?.region).toBeUndefined();
    expect(report?.uncachedFormulaCells).toBe(1);
    expect(
      await readWorkbookTablesBytes({ name: "summary.xlsx", bytes }),
    ).toEqual([]);
  });

  it("counts nothing for a worksheet of ordinary values", async () => {
    const directory = await createTemporaryDirectory();
    const path_ = path.join(directory, "north.xlsx");
    const bytes = await writeWorkbook(path_, [REVIEW_LOG]);

    const [report] = await readWorkbookWorksheetsBytes({
      name: "north.xlsx",
      bytes,
    });
    expect(report?.uncachedFormulaCells).toBe(0);
  });

  it("leaves the table list exactly as it was", async () => {
    const bytes = (
      await input(
        `<row r="1">${textCell("A1", "Case_ID")}${textCell("B1", "Region")}</row>` +
          `<row r="2">${textCell("A2", "R-1")}${textCell("B2", "north")}</row>`,
      )
    ).bytes;

    const reports = await readWorkbookWorksheetsBytes({
      name: "north.xlsx",
      bytes,
    });
    const tables = await readWorkbookTablesBytes({ name: "north.xlsx", bytes });

    // One walk seen two ways, so a table can never differ between them.
    expect(tables).toEqual(
      reports.map((report) => report.table).filter((table) => table),
    );
  });
});

describe("worksheets holding error values", () => {
  const input = async (rows: string) => ({
    name: "north.xlsx",
    bytes: await uncalculatedWorkbookBytes(rows),
  });

  it("counts a typed error and a failed formula alike", async () => {
    const [report] = await readWorkbookWorksheetsBytes(
      await input(
        `<row r="1">${textCell("A1", "Customer")}${textCell("B1", "Amount")}</row>` +
          `<row r="2">${textCell("A2", "Acme")}<c r="B2" t="e"><v>#REF!</v></c></row>` +
          `<row r="3">${textCell("A3", "Beta")}<c r="B3" t="e"><f>1/0</f><v>#DIV/0!</v></c></row>`,
      ),
    );

    expect(report?.errorCells).toBe(2);
    // And this is why the count has to travel beside the table: the engine
    // hands back the internal code Excel numbers each error by, so the amounts
    // read as ordinary numbers that were never in the worksheet.
    expect(report?.table?.rows).toEqual([
      { Customer: "Acme", Amount: 23 },
      { Customer: "Beta", Amount: 7 },
    ]);
  });

  it("counts an error in the header row", async () => {
    const [report] = await readWorkbookWorksheetsBytes(
      await input(
        `<row r="1">${textCell("A1", "Customer")}<c r="B1" t="e"><v>#NAME?</v></c></row>` +
          `<row r="2">${textCell("A2", "Acme")}<c r="B2"><v>5</v></c></row>`,
      ),
    );

    expect(report?.errorCells).toBe(1);
  });

  it("leaves an error above the header alone", async () => {
    // An error cell reads as a value, so it never pushes the header row down by
    // itself; the header is declared here instead, which is what a caller does
    // when a worksheet opens with a title block. The cell above it is outside
    // the rectangle the rows were read from, and counting it would refuse the
    // worksheet over a cell the import never reads - the same region rule the
    // uncalculated formulas follow.
    const rows =
      `<row r="1"><c r="A1" t="e"><v>#REF!</v></c></row>` +
      `<row r="2">${textCell("A2", "Customer")}${textCell("B2", "Region")}</row>` +
      `<row r="3">${textCell("A3", "Acme")}${textCell("B3", "north")}</row>`;

    const [inside] = await readWorkbookWorksheetsBytes(await input(rows));
    expect(inside?.region?.headerRow).toBe(1);
    expect(inside?.errorCells).toBe(1);

    const [outside] = await readWorkbookWorksheetsBytes(await input(rows), {
      headerRow: 2,
    });
    expect(outside?.region?.headerRow).toBe(2);
    expect(outside?.errorCells).toBe(0);
    expect(outside?.table?.rows).toEqual([
      { Customer: "Acme", Region: "north" },
    ]);
  });

  it("counts the two conditions separately in one read", async () => {
    const [report] = await readWorkbookWorksheetsBytes(
      await input(
        `<row r="1">${textCell("A1", "Customer")}${textCell("B1", "Amount")}${textCell("C1", "Share")}</row>` +
          `<row r="2">${textCell("A2", "Acme")}<c r="B2" t="e"><v>#REF!</v></c><c r="C2"><f>B2/10</f></c></row>`,
      ),
    );

    expect(report?.errorCells).toBe(1);
    expect(report?.uncachedFormulaCells).toBe(1);
  });

  it("counts nothing for a worksheet of ordinary values", async () => {
    const directory = await createTemporaryDirectory();
    const path_ = path.join(directory, "north.xlsx");
    const bytes = await writeWorkbook(path_, [REVIEW_LOG]);

    const [report] = await readWorkbookWorksheetsBytes({
      name: "north.xlsx",
      bytes,
    });
    expect(report?.errorCells).toBe(0);
  });
});

describe("worksheets the model cannot parse", () => {
  // The package opens in two steps and only the first is eager: the worksheet
  // part is parsed the first time somebody asks for it, so a malformed row
  // surfaces from a property access rather than from the load.
  const malformed =
    `<row r="1">${textCell("A1", "Customer")}</row>` +
    `<row r="0">${textCell("A2", "Acme")}</row>`;

  it("reports a worksheet read as the stable read failure, naming the sheet", async () => {
    const bytes = await uncalculatedWorkbookBytes(malformed);

    await expect(
      readWorkbookWorksheetsBytes({ name: "north.xlsx", bytes }),
    ).rejects.toMatchObject({
      code: XLSX_ERRORS.XLSX_READ_FAILED,
      details: { source: "north.xlsx", worksheet: "Review Log" },
    });
  });

  it("reports the same failure from the description of the same file", async () => {
    const bytes = await uncalculatedWorkbookBytes(malformed);

    await expect(
      describeWorkbookBytes({ name: "north.xlsx", bytes }),
    ).rejects.toMatchObject({
      code: XLSX_ERRORS.XLSX_READ_FAILED,
      details: { source: "north.xlsx", worksheet: "Review Log" },
    });
  });
});

describe("worksheets that leave positions to document order", () => {
  // The r attribute is optional on both <row> and <c>: a worksheet may rely on
  // document order instead. A generator that writes one is a workbook a person
  // can open, so refusing it because the model wanted an attribute the format
  // makes optional refuses a readable file.
  const implicit =
    `<row><c t="inlineStr"><is><t>Case_ID</t></is></c><c t="inlineStr"><is><t>Region</t></is></c></row>` +
    `<row><c t="inlineStr"><is><t>R-1</t></is></c><c t="inlineStr"><is><t>north</t></is></c></row>` +
    `<row><c t="inlineStr"><is><t>R-2</t></is></c><c><f>1+1</f></c></row>`;

  it("reads rows and cells that carry no reference", async () => {
    const bytes = await uncalculatedWorkbookBytes(implicit);

    const [report] = await readWorkbookWorksheetsBytes({
      name: "implicit.xlsx",
      bytes,
    });

    // Document order put the header on row 1 and the values under it.
    expect(report?.region?.headerRow).toBe(1);
    expect(report?.table?.rows).toEqual([
      { Case_ID: "R-1", Region: "north" },
      { Case_ID: "R-2", Region: null },
    ]);
    // And the formula in the last cell was found, which is the whole reason the
    // model has to be able to read this worksheet at all.
    expect(report?.uncachedFormulaCells).toBe(1);
  });

  it("describes such a worksheet through the L1 consumers", async () => {
    const bytes = await uncalculatedWorkbookBytes(implicit);

    const { description } = await describeWorkbookBytes({
      name: "implicit.xlsx",
      bytes,
    });
    const sheet = description.sheets[0];
    expect(sheet?.headerRow).toBe(1);
    expect(sheet?.rowCount).toBe(3);
    expect(sheet?.dataRowCount).toBe(2);
    expect(sheet?.columns.map((column) => column.header)).toEqual([
      "Case_ID",
      "Region",
    ]);
  });

  it("still refuses a reference that is present and unreadable", async () => {
    // Absent is the format saying "use document order". Present and malformed
    // is a damaged file, and stays an error rather than being guessed at. The
    // rows are parsed on demand rather than at load, so this arrives from a
    // property access; it is translated into the same read failure the load
    // itself raises, and the parser's own message stays reachable as the cause.
    const bytes = await uncalculatedWorkbookBytes(
      `<row r="1"><c r="not-a-cell" t="inlineStr"><is><t>Case_ID</t></is></c></row>`,
    );

    const failure = await describeWorkbookBytes({
      name: "broken.xlsx",
      bytes,
    }).catch((error: unknown) => error);

    expect(failure).toMatchObject({
      code: XLSX_ERRORS.XLSX_READ_FAILED,
      details: { source: "broken.xlsx", worksheet: "Review Log" },
    });
    expect((failure as { cause?: unknown }).cause).toMatchObject({
      message: expect.stringContaining("invalid cell reference: not-a-cell"),
    });
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

  it("reports the same worksheets from a path as from bytes", async () => {
    // Both surfaces adapt their input and hand the same operation the same
    // bytes, so neither can answer differently about the same workbook. The
    // formula count is part of that answer.
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "north.xlsx");
    const bytes = await uncalculatedWorkbookBytes(
      `<row r="1">${textCell("A1", "Case_ID")}${textCell("B1", "Failed Checks")}</row>` +
        `<row r="2">${textCell("A2", "R-1")}<c r="B2"><f>1+1</f></c></row>`,
    );
    await writeFile(input, bytes);

    const fromBytes = await readWorkbookWorksheetsBytes({
      name: "north.xlsx",
      bytes,
    });
    const fromFile = await readWorkbookWorksheets(input);

    expect(fromFile).toEqual(fromBytes);
    expect(fromFile[0]?.uncachedFormulaCells).toBe(1);
    expect(fromFile[0]?.region?.headerRow).toBe(1);
  });

  it("reads the same worksheet tables from bytes as from a path", async () => {
    const directory = await createTemporaryDirectory();
    const input = path.join(directory, "north.xlsx");
    const bytes = await writeWorkbook(input, [
      REVIEW_LOG,
      { hidden: 1, name: "Notes", rows: [["Note"], ["kept back"]] },
    ]);

    const fromFile = await readWorkbookTables(input);
    const fromBytes = await readWorkbookTablesBytes({
      name: "north.xlsx",
      bytes,
    });

    expect(fromBytes).toEqual(fromFile);
    expect(fromBytes).toHaveLength(1);
    expect(fromBytes[0]).toMatchObject({
      columns: ["Case_ID", "Region", "Failed Checks"],
      rows: [
        { Case_ID: "R-1", Region: "north", "Failed Checks": 5 },
        { Case_ID: "R-2", Region: "south", "Failed Checks": 7 },
        { Case_ID: "R-3", Region: "north", "Failed Checks": 9 },
      ],
    });

    // The selection options travel with the reader, hidden worksheets included.
    expect(
      (
        await readWorkbookTablesBytes(
          { name: "north.xlsx", bytes },
          { includeHiddenSheets: true },
        )
      ).map((table) => table.source?.sheet),
    ).toEqual(["Review Log", "Notes"]);
    expect(
      await readWorkbookTablesBytes(
        { name: "north.xlsx", bytes },
        { sheets: ["missing"] },
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
