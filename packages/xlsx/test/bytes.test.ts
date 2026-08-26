import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";

import {
  isConsultChimpsError,
  OPERATION_ABORTED,
  type OperationProgress,
} from "@consultchimps/core";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import * as XLSX from "xlsx";

import {
  mergeWorkbooksBytes,
  planSplitWorkbookBytes,
  readWorksheetRecordsBytes,
  splitWorkbookBytes,
} from "../src/bytes.js";

type CellInput = boolean | null | number | string;

const structuredTableFixture = fileURLToPath(
  new URL("./fixtures/structured-table.xlsx", import.meta.url),
);
// A DOS timestamp has two-second resolution, so a shorter pause could hide a
// clock-dependent byte difference inside one tick.
const CLOCK_TICK = 2_100;

function workbookBytes(
  sheets: Array<[string, CellInput[][]]>,
  options: {
    names?: Array<{ Name: string; Ref: string }>;
    visibility?: Record<string, 0 | 1 | 2>;
  } = {},
): Uint8Array {
  const workbook = XLSX.utils.book_new();
  for (const [sheetName, rows] of sheets) {
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet(rows),
      sheetName,
    );
  }
  if (options.names) {
    workbook.Workbook = { ...workbook.Workbook, Names: options.names };
  }
  if (options.visibility) {
    workbook.Workbook = {
      ...workbook.Workbook,
      Sheets: sheets.map(([sheetName]) => ({
        Hidden: options.visibility?.[sheetName] ?? 0,
        name: sheetName,
      })),
    };
  }
  return new Uint8Array(
    XLSX.write(workbook, {
      bookType: "xlsx",
      cellStyles: true,
      compression: true,
      type: "array",
    }) as ArrayBuffer,
  );
}

/**
 * A workbook package whose ZIP header is intact but whose contents are not,
 * which is what a truncated or damaged .xlsx upload looks like. Arbitrary
 * bytes without a ZIP header are not an error: the workbook reader treats
 * them as plain text, exactly as the path-based operations do.
 */
function corruptWorkbookBytes(): Uint8Array {
  return new Uint8Array([0x50, 0x4b, 0x03, 0x04, 9, 9, 9, 9, 9, 9, 9, 9]);
}

function clientRows(): Array<[string, CellInput[][]]> {
  return [
    [
      "Clients",
      [
        ["Client", "Region", "Amount"],
        ["A", "North", 10],
        ["B", "South", 20],
        ["C", "North", 30],
      ],
    ],
  ];
}

function sheetNames(bytes: Uint8Array): string[] {
  return XLSX.read(bytes, { type: "array" }).SheetNames;
}

function readSheet(
  bytes: Uint8Array,
  sheetName: string,
): Array<Record<string, unknown>> {
  const workbook = XLSX.read(bytes, { type: "array" });
  return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]!, {
    defval: null,
    raw: true,
  });
}

const WORKBOOK_MAIN_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const MACRO_WORKBOOK_MAIN_CONTENT_TYPE =
  "application/vnd.ms-excel.sheet.macroEnabled.main+xml";

/**
 * A macro-enabled package: the VBA project, the `bin` default that types it,
 * and - the part that makes the package coherent rather than merely
 * suggestive - the main workbook part declared as macro-enabled. A package
 * carrying `vbaProject.bin` while still declaring itself an ordinary workbook
 * is the contradiction the split refuses, so a fixture meant to be valid has
 * to change both.
 */
async function macroWorkbookBytes(
  options: { declareMacroContentType?: boolean } = {},
): Promise<Uint8Array> {
  const archive = await JSZip.loadAsync(workbookBytes(clientRows()));
  archive.file("xl/vbaProject.bin", Buffer.from([0xd0, 0xcf, 0x11, 0xe0]));
  const contentTypes = await archive.file("[Content_Types].xml")!.async("text");
  const withDefault = contentTypes.replace(
    "</Types>",
    '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>',
  );
  const declared =
    options.declareMacroContentType === false
      ? withDefault
      : withDefault.replace(
          WORKBOOK_MAIN_CONTENT_TYPE,
          MACRO_WORKBOOK_MAIN_CONTENT_TYPE,
        );
  expect(declared).toContain("vbaProject");
  archive.file("[Content_Types].xml", declared);
  return archive.generateAsync({ compression: "DEFLATE", type: "uint8array" });
}

async function structuredTableBytes(hidden = false): Promise<Uint8Array> {
  const bytes = new Uint8Array(await readFile(structuredTableFixture));
  if (!hidden) {
    return bytes;
  }

  const archive = await JSZip.loadAsync(bytes);
  const workbookXml = await archive.file("xl/workbook.xml")!.async("text");
  const hiddenWorkbookXml = workbookXml.replace(
    'name="Clients" sheetId="2"',
    'name="Clients" state="hidden" sheetId="2"',
  );
  expect(hiddenWorkbookXml).not.toBe(workbookXml);
  archive.file("xl/workbook.xml", hiddenWorkbookXml);
  return archive.generateAsync({ compression: "DEFLATE", type: "uint8array" });
}

describe("byte-level workbook splitting", () => {
  it("splits worksheet rows in memory and reports names instead of paths", async () => {
    const events: OperationProgress[] = [];
    const { result, outputs } = await splitWorkbookBytes({
      input: { name: "client list.xlsx", bytes: workbookBytes(clientRows()) },
      column: " region ",
      onProgress: (progress) => events.push(progress),
    });

    expect(result.operation).toBe("sheets.split-by-column");
    expect(result.metrics).toEqual({
      calcChainEntriesRemoved: 0,
      formulaCellsBlankedForRemovedRows: 0,
      formulaCellsConverted: 0,
      formulaCellsWithoutCachedValues: 0,
      groups: 2,
      inputFiles: 1,
      inputRows: 3,
      outputFiles: 2,
      outputRows: 3,
      pivotTablesRemoved: 0,
      rowsDeleted: 3,
      sheetsCopiedUnchanged: 0,
      sheetsFiltered: 1,
      skippedRows: 0,
      valuesOnly: 0,
    });
    expect(outputs.map((output) => output.name)).toEqual([
      "client list-North.xlsx",
      "client list-South.xlsx",
    ]);
    expect(result.artifacts.map((artifact) => artifact.path)).toEqual(
      outputs.map((output) => output.name),
    );
    expect(
      outputs.every(
        (output) =>
          output.mediaType ===
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      ),
    ).toBe(true);
    expect(events.map((event) => [event.stage, event.completed])).toEqual([
      ["building-workbooks", 1],
      ["building-workbooks", 2],
    ]);

    expect(readSheet(outputs[0]!.bytes, "Clients")).toEqual([
      { Amount: 10, Client: "A", Region: "North" },
      { Amount: 30, Client: "C", Region: "North" },
    ]);
    expect(readSheet(outputs[1]!.bytes, "Clients")).toEqual([
      { Amount: 20, Client: "B", Region: "South" },
    ]);

    // The all-worksheet default reports which worksheets it filtered and how
    // many rows each output kept, the same detail the file surface reports.
    expect(result.summary).toEqual({
      column: " region ",
      copiedUnchangedSheets: [],
      filteredSheets: ["Clients"],
      input: "client list.xlsx",
      valuesOnly: false,
    });
    expect(result.outputs).toEqual([
      {
        formulaCellsConverted: 0,
        formulaCellsWithoutCachedValues: 0,
        output: "client list-North.xlsx",
        sheets: [{ deletedRows: 1, retainedRows: 2, sheet: "Clients" }],
        value: "North",
      },
      {
        formulaCellsConverted: 0,
        formulaCellsWithoutCachedValues: 0,
        output: "client list-South.xlsx",
        sheets: [{ deletedRows: 2, retainedRows: 1, sheet: "Clients" }],
        value: "South",
      },
    ]);
  });

  it("keeps every worksheet, filtered or not, in each output", async () => {
    const input = {
      name: "regions.xlsx",
      bytes: workbookBytes([
        [
          "Current",
          [
            ["Client", "Region"],
            ["A", "North"],
            ["B", "South"],
          ],
        ],
        [
          "Archive",
          [
            ["Client", "Region"],
            ["C", "South"],
          ],
        ],
        [
          "Notes",
          [
            ["Measure", "Value"],
            ["Total", 99],
          ],
        ],
      ]),
    };

    const { outputs, result } = await splitWorkbookBytes({
      input,
      column: "Region",
    });

    // Every worksheet survives in every output: the two that carry the column
    // keep only their group's rows, and the one that does not is copied
    // through untouched. The compact single-source split cannot do this - it
    // writes one worksheet built from the values it read.
    expect(outputs.map((output) => output.name)).toEqual([
      "regions-North.xlsx",
      "regions-South.xlsx",
    ]);
    for (const output of outputs) {
      expect(sheetNames(output.bytes)).toEqual(["Current", "Archive", "Notes"]);
      expect(readSheet(output.bytes, "Notes")).toEqual([
        { Measure: "Total", Value: 99 },
      ]);
    }
    expect(readSheet(outputs[0]!.bytes, "Current")).toEqual([
      { Client: "A", Region: "North" },
    ]);
    expect(readSheet(outputs[0]!.bytes, "Archive")).toEqual([]);
    expect(readSheet(outputs[1]!.bytes, "Archive")).toEqual([
      { Client: "C", Region: "South" },
    ]);
    expect(result.summary).toMatchObject({
      copiedUnchangedSheets: ["Notes"],
      filteredSheets: ["Current", "Archive"],
    });
    expect(result.metrics).toMatchObject({
      groups: 2,
      inputRows: 3,
      sheetsCopiedUnchanged: 1,
      sheetsFiltered: 2,
    });
  });

  it("still rebuilds one compact worksheet when asked not to preserve", async () => {
    const input = {
      name: "regions.xlsx",
      bytes: workbookBytes([
        [
          "Clients",
          [
            ["Client", "Region"],
            ["A", "North"],
            ["B", "South"],
          ],
        ],
        [
          "Notes",
          [
            ["Measure", "Value"],
            ["Total", 99],
          ],
        ],
      ]),
    };

    const { outputs, result } = await splitWorkbookBytes({
      input,
      column: "Region",
      preserveWorkbook: false,
      sheet: "Clients",
    });

    expect(sheetNames(outputs[0]!.bytes)).toEqual(["Clients"]);
    expect(result.metrics).toMatchObject({
      groups: 2,
      sheetsCopiedUnchanged: 0,
      sheetsFiltered: 1,
    });
    // The narrower modes report no all-worksheet detail, so a caller can tell
    // which engine ran without inspecting the outputs.
    expect(result.outputs).toBeUndefined();
    expect(result.summary).toBeUndefined();
  });

  it("compares values strictly on request", async () => {
    const input = {
      name: "regions.xlsx",
      bytes: workbookBytes([
        [
          "Clients",
          [
            ["Client", "Region"],
            ["A", "North"],
            ["B", " north "],
            ["C", "NORTH"],
          ],
        ],
      ]),
    };

    // Default matching folds case and surrounding whitespace, so the three
    // spellings are one group; strict matching keeps them apart.
    const tolerant = await splitWorkbookBytes({ input, column: "Region" });
    expect(tolerant.result.metrics.groups).toBe(1);
    expect(tolerant.outputs.map((output) => output.name)).toEqual([
      "regions-North.xlsx",
    ]);

    const strict = await splitWorkbookBytes({
      input,
      column: "Region",
      strict: true,
    });
    expect(strict.result.metrics.groups).toBe(3);
    expect(strict.result.outputs?.map((output) => output.value)).toEqual([
      "North",
      "north",
      "NORTH",
    ]);
    // Three values that sanitize to one filename are told apart by a stable
    // suffix, compared case-insensitively because Windows would treat the
    // unsuffixed names as one file.
    expect(new Set(strict.outputs.map((output) => output.name)).size).toBe(3);
  });

  it("plans a split, including skipped rows, without producing any bytes", async () => {
    const input = {
      name: "clients.xlsx",
      bytes: workbookBytes([
        [
          "Clients",
          [
            ["Client", "Region"],
            ["A", "North"],
            ["B", null],
            ["C", "South"],
          ],
        ],
      ]),
    };

    const plan = await planSplitWorkbookBytes({
      input,
      column: "Region",
      includeBlank: false,
    });
    expect(plan.operation).toBe("sheets.split-by-column");
    expect(plan.inputs).toEqual(["clients.xlsx"]);
    expect(plan.outputs.map((output) => output.path)).toEqual([
      "clients-North.xlsx",
      "clients-South.xlsx",
    ]);
    expect(plan.outputs.every((output) => output.exists === false)).toBe(true);
    expect(plan.warnings).toEqual([
      'Skipped 1 row with blank values in "Region"; no blank-value workbook was created.',
    ]);
    expect(plan.metrics).toEqual({
      calcChainEntriesRemoved: 0,
      formulaCellsBlankedForRemovedRows: 0,
      formulaCellsConverted: 0,
      formulaCellsWithoutCachedValues: 0,
      groups: 2,
      inputFiles: 1,
      inputRows: 3,
      outputFiles: 2,
      pivotTablesRemoved: 0,
      rowsDeleted: 0,
      sheetsCopiedUnchanged: 0,
      sheetsFiltered: 1,
      skippedRows: 1,
      valuesOnly: 0,
    });

    const executed = await splitWorkbookBytes({
      input,
      column: "Region",
      includeBlank: false,
    });
    expect(executed.result.warnings).toEqual(plan.warnings);
    expect(executed.result.metrics.skippedRows).toBe(1);
    // A plan reads the source but builds nothing, so the metrics that describe
    // reading match the split's and the metrics that describe writing are zero
    // until the workbooks exist.
    expect(plan.metrics.rowsDeleted).toBe(0);
    // Every output drops the rows that are not its own, so the count is the
    // sum across outputs rather than a count of source rows.
    expect(executed.result.metrics.rowsDeleted).toBe(4);
  });

  it("keeps the whole workbook when splitting a named Excel Table", async () => {
    const bytes = await structuredTableBytes();
    const { outputs } = await splitWorkbookBytes({
      input: { name: "clients.xlsx", bytes },
      column: "Region",
      table: "ClientData",
    });

    expect(outputs.map((output) => output.name)).toEqual([
      "clients-North.xlsx",
      "clients-South.xlsx",
    ]);

    const sourceArchive = await JSZip.loadAsync(bytes);
    const northArchive = await JSZip.loadAsync(outputs[0]!.bytes);
    for (const part of [
      "xl/styles.xml",
      "xl/theme/theme1.xml",
      "xl/workbook.xml",
      "xl/worksheets/sheet1.xml",
    ]) {
      expect(await northArchive.file(part)?.async("text")).toBe(
        await sourceArchive.file(part)?.async("text"),
      );
    }

    const northWorksheetXml = await northArchive
      .file("xl/worksheets/sheet2.xml")!
      .async("text");
    expect(northWorksheetXml).toContain("Client allocation report");
    expect(northWorksheetXml).toContain('<x:mergeCell ref="A1:D2" />');
    expect(
      await northArchive.file("xl/tables/table1.xml")!.async("text"),
    ).toContain('ref="B4:D7"');

    const compact = await splitWorkbookBytes({
      input: { name: "clients.xlsx", bytes },
      column: "Region",
      preserveWorkbook: false,
      table: "clientdata",
    });
    expect(readSheet(compact.outputs[0]!.bytes, "Clients")).toEqual([
      { Amount: 10, Client: "A", Region: "North" },
      { Amount: 30, Client: "C", Region: "North" },
    ]);
  });

  it("replaces formulas with cached values in a preserved split", async () => {
    const source = await JSZip.loadAsync(await structuredTableBytes());
    const sheetXml = await source
      .file("xl/worksheets/sheet2.xml")!
      .async("text");
    source.file(
      "xl/worksheets/sheet2.xml",
      sheetXml.replace(
        '<x:c r="D7" s="6" t="n"><x:v>30</x:v></x:c>',
        '<x:c r="D7" s="6" t="n"><x:f>ClientData[[#This Row],[Amount]]</x:f><x:v>30</x:v></x:c>',
      ),
    );
    const bytes = await source.generateAsync({
      compression: "DEFLATE",
      type: "uint8array",
    });

    const { outputs } = await splitWorkbookBytes({
      input: { name: "clients.xlsx", bytes },
      column: "Region",
      table: "ClientData",
      values: true,
    });
    const northArchive = await JSZip.loadAsync(outputs[0]!.bytes);
    const northSheetXml = await northArchive
      .file("xl/worksheets/sheet2.xml")!
      .async("text");
    expect(northSheetXml).not.toContain("<x:f>");
    expect(northSheetXml).toContain(
      '<x:c r="D6" s="6" t="n"><x:v>30</x:v></x:c>',
    );
  });

  it("splits a named range and requires the matching source selection", async () => {
    const bytes = workbookBytes(
      [
        [
          "Clients",
          [
            ["Quarterly report", null, null],
            ["Client", "Region", "Amount"],
            ["A", "North", 10],
            ["B", "South", 20],
            ["Total", null, 30],
          ],
        ],
      ],
      { names: [{ Name: "ClientRange", Ref: "Clients!$A$2:$C$4" }] },
    );
    const input = { name: "clients.xlsx", bytes };

    const { outputs, result } = await splitWorkbookBytes({
      input,
      column: "Region",
      range: "clientrange",
    });
    expect(result.metrics.inputRows).toBe(2);
    expect(outputs.map((output) => output.name)).toEqual([
      "clients-North.xlsx",
      "clients-South.xlsx",
    ]);
    expect(readSheet(outputs[0]!.bytes, "Clients")).toEqual([
      { Amount: 10, Client: "A", Region: "North" },
    ]);

    await expect(
      splitWorkbookBytes({
        input,
        column: "Region",
        headerRow: 2,
        range: "ClientRange",
      }),
    ).rejects.toMatchObject({ code: "XLSX_SPLIT_RANGE_HEADER_ROW" });
    await expect(
      splitWorkbookBytes({
        input,
        column: "Region",
        range: "ClientRange",
        table: "ClientData",
      }),
    ).rejects.toMatchObject({ code: "XLSX_SPLIT_TABLE_RANGE_CONFLICT" });
    // Whole-workbook preservation is the default with no source named, so the
    // refusal now only reaches a caller who named a source that cannot offer
    // it - a named range, or a worksheet.
    await expect(
      splitWorkbookBytes({
        input,
        column: "Region",
        preserveWorkbook: true,
        range: "ClientRange",
      }),
    ).rejects.toMatchObject({ code: "XLSX_SPLIT_PRESERVE_REQUIRES_TABLE" });
    await expect(
      splitWorkbookBytes({ input, column: "Region", range: "Missing" }),
    ).rejects.toMatchObject({ code: "XLSX_SPLIT_NO_TABLE" });
  });

  it("honours worksheet, header row, and hidden sheet selection", async () => {
    const input = {
      name: "regions.xlsx",
      bytes: workbookBytes([
        [
          "Current",
          [
            ["Ignored title", null],
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
      ]),
    };

    // Two worksheets carrying the column are what the all-worksheet default
    // exists for; it filters both. Only the compact rebuild, which writes a
    // single worksheet, still has to be told which one to read.
    const both = await splitWorkbookBytes({ input, column: "Region" });
    expect(both.result.metrics).toMatchObject({
      groups: 2,
      sheetsFiltered: 2,
    });
    await expect(
      splitWorkbookBytes({ input, column: "Region", preserveWorkbook: false }),
    ).rejects.toMatchObject({ code: "XLSX_SPLIT_MULTIPLE_TABLES" });

    const selected = await splitWorkbookBytes({
      input,
      column: "Region",
      headerRow: 2,
      sheet: "current",
    });
    expect(readSheet(selected.outputs[0]!.bytes, "Current")).toEqual([
      { Client: "A", Region: "North" },
    ]);
    await expect(
      splitWorkbookBytes({ input, column: "Region", headerRow: 0 }),
    ).rejects.toMatchObject({ code: "XLSX_INVALID_HEADER_ROW" });

    const hidden = {
      name: "clients.xlsx",
      bytes: await structuredTableBytes(true),
    };
    await expect(
      splitWorkbookBytes({
        input: hidden,
        column: "Region",
        table: "ClientData",
      }),
    ).rejects.toMatchObject({ code: "XLSX_SPLIT_NO_TABLE" });
    const included = await splitWorkbookBytes({
      input: hidden,
      column: "Region",
      includeHiddenSheets: true,
      table: "ClientData",
    });
    expect(included.result.metrics.outputFiles).toBe(2);
  });

  it("sanitizes output names and disambiguates repeated group values", async () => {
    const bytes = workbookBytes([
      [
        "Data",
        [
          ["Region", "Amount"],
          ["North/West", 1],
          ["North:West", 2],
          [null, 3],
        ],
      ],
    ]);

    const { outputs } = await splitWorkbookBytes({
      input: { name: "cli*ent<>report.xlsx", bytes },
      column: "Region",
    });
    // The all-worksheet split never writes a workbook for a blank key, so the
    // third row is skipped rather than named "blank".
    expect(outputs.map((output) => output.name)).toEqual([
      "cli-ent-report-North-West.xlsx",
      "cli-ent-report-North-West-2.xlsx",
    ]);

    const blankIncluded = await splitWorkbookBytes({
      input: { name: "cli*ent<>report.xlsx", bytes },
      column: "Region",
      includeBlank: true,
      preserveWorkbook: false,
      sheet: "Data",
    });
    expect(blankIncluded.outputs.at(-1)?.name).toBe(
      "cli-ent-report-blank.xlsx",
    );

    const reserved = await splitWorkbookBytes({
      input: { name: "CON.xlsx", bytes },
      column: "Region",
    });
    expect(reserved.outputs[0]?.name).toBe("_CON-North-West.xlsx");

    const prefixed = await splitWorkbookBytes({
      input: { name: "report.xlsx", bytes },
      column: "Region",
      filenamePrefix: "  ",
    });
    expect(prefixed.outputs[0]?.name).toBe("split-North-West.xlsx");

    // Each of these code points encodes to three UTF-8 bytes, so an
    // unbounded name would be far longer than a portable filename allows.
    const longValue = "\u4E2D".repeat(100);
    const bounded = await splitWorkbookBytes({
      input: {
        name: `${"a".repeat(300)}.xlsx`,
        bytes: workbookBytes([
          [
            "Data",
            [
              ["Region", "Amount"],
              [longValue, 1],
            ],
          ],
        ]),
      },
      column: "Region",
    });
    const name = bounded.outputs[0]!.name;
    expect(name.startsWith(`${"a".repeat(80)}-`)).toBe(true);
    expect(Buffer.byteLength(name, "utf8")).toBeLessThanOrEqual(255);
    expect([...name.slice(81, -".xlsx".length)]).toHaveLength(26);
  });

  it("produces byte-identical workbooks for identical inputs", async () => {
    const plain = { name: "clients.xlsx", bytes: workbookBytes(clientRows()) };
    const preserved = {
      name: "clients.xlsx",
      bytes: await structuredTableBytes(),
    };

    const firstAllWorksheets = await splitWorkbookBytes({
      input: plain,
      column: "Region",
    });
    const firstPlain = await splitWorkbookBytes({
      input: plain,
      column: "Region",
      preserveWorkbook: false,
      sheet: "Clients",
    });
    const firstPreserved = await splitWorkbookBytes({
      input: preserved,
      column: "Region",
      table: "ClientData",
    });
    await new Promise((resolve) => setTimeout(resolve, CLOCK_TICK));
    const secondAllWorksheets = await splitWorkbookBytes({
      input: plain,
      column: "Region",
    });
    const secondPlain = await splitWorkbookBytes({
      input: plain,
      column: "Region",
      preserveWorkbook: false,
      sheet: "Clients",
    });
    const secondPreserved = await splitWorkbookBytes({
      input: preserved,
      column: "Region",
      table: "ClientData",
    });

    for (const [first, second] of [
      [firstAllWorksheets, secondAllWorksheets],
      [firstPlain, secondPlain],
      [firstPreserved, secondPreserved],
    ] as const) {
      expect(first.outputs).not.toHaveLength(0);
      for (let index = 0; index < first.outputs.length; index += 1) {
        expect(
          Buffer.compare(
            Buffer.from(first.outputs[index]!.bytes),
            Buffer.from(second.outputs[index]!.bytes),
          ),
        ).toBe(0);
      }
    }
  });

  it("cancels a split without producing partial output", async () => {
    const input = { name: "clients.xlsx", bytes: workbookBytes(clientRows()) };

    const beforeStart = new AbortController();
    beforeStart.abort();
    let thrown: unknown;
    try {
      await splitWorkbookBytes({
        input,
        column: "Region",
        signal: beforeStart.signal,
      });
    } catch (error) {
      thrown = error;
    }
    expect(isConsultChimpsError(thrown)).toBe(true);
    expect((thrown as { code: string }).code).toBe(OPERATION_ABORTED);
    expect((thrown as Error).message).toContain(
      "no partial output was produced",
    );
    expect((thrown as Error).message).not.toContain("output files");

    // Cancelling between groups stops the next workbook from being built.
    const midway = new AbortController();
    const stages: number[] = [];
    await expect(
      splitWorkbookBytes({
        input,
        column: "Region",
        signal: midway.signal,
        onProgress: (progress) => {
          stages.push(progress.completed);
          midway.abort();
        },
      }),
    ).rejects.toMatchObject({ code: OPERATION_ABORTED });
    expect(stages).toEqual([1]);

    // Cancelling while the final workbook is serialized still returns nothing.
    const atEnd = new AbortController();
    await expect(
      splitWorkbookBytes({
        input: {
          name: "single.xlsx",
          bytes: workbookBytes([
            [
              "Data",
              [
                ["Region", "Amount"],
                ["North", 1],
              ],
            ],
          ]),
        },
        column: "Region",
        signal: atEnd.signal,
        onProgress: () => atEnd.abort(),
      }),
    ).rejects.toMatchObject({ code: OPERATION_ABORTED });
  });

  it("reports unreadable workbooks and empty selections", async () => {
    await expect(
      splitWorkbookBytes({
        input: { name: "corrupt.xlsx", bytes: corruptWorkbookBytes() },
        column: "Region",
      }),
    ).rejects.toMatchObject({ code: "XLSX_READ_FAILED" });

    await expect(
      splitWorkbookBytes({
        input: {
          name: "clients.xlsx",
          bytes: workbookBytes([
            [
              "Clients",
              [
                ["Client", "Region"],
                ["A", null],
              ],
            ],
          ]),
        },
        column: "Region",
        includeBlank: false,
      }),
    ).rejects.toMatchObject({ code: "XLSX_SPLIT_NO_GROUPS" });

    // No worksheet carries the column, which the all-worksheet split reports
    // as the missing column rather than as a missing source: it never asked
    // for one source in particular.
    await expect(
      splitWorkbookBytes({
        input: { name: "empty.xlsx", bytes: workbookBytes([["Data", []]]) },
        column: "Region",
      }),
    ).rejects.toMatchObject({ code: "XLSX_SPLIT_COLUMN_NOT_FOUND" });
    await expect(
      splitWorkbookBytes({
        input: { name: "empty.xlsx", bytes: workbookBytes([["Data", []]]) },
        column: "Region",
        preserveWorkbook: false,
      }),
    ).rejects.toMatchObject({ code: "XLSX_SPLIT_NO_TABLE" });
  });

  it("refuses a name that is not an Excel workbook before reading it", async () => {
    // The name decides the outputs' extension as well as whether the split can
    // open the input at all, so it is checked before the bytes are parsed.
    await expect(
      splitWorkbookBytes({
        input: { name: "clients.csv", bytes: workbookBytes(clientRows()) },
        column: "Region",
      }),
    ).rejects.toMatchObject({ code: "XLSX_SPLIT_UNSUPPORTED_FILE" });
    await expect(
      planSplitWorkbookBytes({
        input: { name: "clients", bytes: workbookBytes(clientRows()) },
        column: "Region",
      }),
    ).rejects.toMatchObject({ code: "XLSX_SPLIT_UNSUPPORTED_FILE" });
  });

  it("refuses a workbook whose package contradicts its name", async () => {
    // Ordinary workbook bytes offered under an .xlsm name. Trusting the name
    // would advertise every output as macro-enabled while its package still
    // declared an ordinary workbook, which is what Excel warns about.
    const plainUnderMacroName = splitWorkbookBytes({
      input: { name: "clients.xlsm", bytes: workbookBytes(clientRows()) },
      column: "Region",
    });
    await expect(plainUnderMacroName).rejects.toMatchObject({
      code: "XLSX_SPLIT_PACKAGE_TYPE_MISMATCH",
      details: {
        declaredExtension: ".xlsx",
        macroEnabled: false,
        nameExtension: ".xlsm",
      },
    });

    // The inverse: a macro-enabled package offered under an .xlsx name would
    // otherwise put a macro project inside a file whose name denies it.
    await expect(
      splitWorkbookBytes({
        input: { name: "clients.xlsx", bytes: await macroWorkbookBytes() },
        column: "Region",
      }),
    ).rejects.toMatchObject({
      code: "XLSX_SPLIT_PACKAGE_TYPE_MISMATCH",
      details: { declaredExtension: ".xlsm", macroEnabled: true },
    });

    // A macro project alone does not make a package macro-enabled; the main
    // workbook part's declared content type does. This one still says .xlsx.
    await expect(
      splitWorkbookBytes({
        input: {
          name: "clients.xlsm",
          bytes: await macroWorkbookBytes({ declareMacroContentType: false }),
        },
        column: "Region",
      }),
    ).rejects.toMatchObject({ code: "XLSX_SPLIT_PACKAGE_TYPE_MISMATCH" });

    // The preview refuses for the same reason, so it can never promise a split
    // that the run would then refuse.
    await expect(
      planSplitWorkbookBytes({
        input: { name: "clients.xlsm", bytes: workbookBytes(clientRows()) },
        column: "Region",
      }),
    ).rejects.toMatchObject({ code: "XLSX_SPLIT_PACKAGE_TYPE_MISMATCH" });
  });

  it("keeps a macro workbook macro-enabled", async () => {
    const { outputs, result } = await splitWorkbookBytes({
      input: { name: "clients.xlsm", bytes: await macroWorkbookBytes() },
      column: "Region",
    });

    expect(outputs.map((output) => output.name)).toEqual([
      "clients-North.xlsm",
      "clients-South.xlsm",
    ]);
    expect(
      outputs.every(
        (output) =>
          output.mediaType === "application/vnd.ms-excel.sheet.macroEnabled.12",
      ),
    ).toBe(true);
    expect(result.artifacts.map((artifact) => artifact.mediaType)).toEqual(
      outputs.map((output) => output.mediaType),
    );
    // The macro project travels with the workbook it belongs to; a package
    // whose content type contradicts its name is one Excel refuses to open.
    for (const output of outputs) {
      const produced = await JSZip.loadAsync(output.bytes);
      expect(produced.file("xl/vbaProject.bin")).not.toBeNull();
    }
  });
});

describe("byte-level workbook merging", () => {
  it("merges every worksheet in order and records provenance", async () => {
    const events: OperationProgress[] = [];
    const { result, outputs } = await mergeWorkbooksBytes({
      inputs: [
        {
          name: "north.xlsx",
          bytes: workbookBytes(
            [
              ["Summary", [["Region"], ["North"]]],
              ["Private", [["Amount"], [100]]],
            ],
            { visibility: { Private: 2 } },
          ),
        },
        {
          name: "south.xlsx",
          bytes: workbookBytes([["Summary", [["Region"], ["South"]]]]),
        },
      ],
      outputName: "client pack.xlsx",
      onProgress: (progress) => events.push(progress),
    });

    expect(result.operation).toBe("sheets.merge");
    expect(result.metrics).toEqual({
      hiddenSheets: 1,
      inputFiles: 2,
      outputSheets: 3,
    });
    expect(result.warnings).toEqual([
      '1 source worksheet was hidden; see the visible "Sheet Index" worksheet.',
    ]);
    expect(events.map((event) => [event.stage, event.completed])).toEqual([
      ["merging-inputs", 1],
      ["merging-inputs", 2],
    ]);
    expect(outputs[0]?.name).toBe("client pack.xlsx");

    const merged = XLSX.read(outputs[0]!.bytes, { type: "array" });
    expect(merged.SheetNames).toEqual([
      "Summary",
      "Private",
      "Summary (2)",
      "Sheet Index",
    ]);
    expect(
      XLSX.utils.sheet_to_json(merged.Sheets["Sheet Index"]!, {
        header: 1,
        raw: true,
      }),
    ).toEqual([
      [
        "Source file",
        "Original worksheet",
        "Final worksheet",
        "Source visibility",
      ],
      ["north.xlsx", "Summary", "Summary", "Visible"],
      ["north.xlsx", "Private", "Private", "Very hidden"],
      ["south.xlsx", "Summary", "Summary (2)", "Visible"],
    ]);
  });

  it("omits the index on request and derives a safe default name", async () => {
    const input = {
      name: "source.xlsx",
      bytes: workbookBytes([["Private", [["Value"], [1]]]], {
        visibility: { Private: 1 },
      }),
    };

    const withoutIndex = await mergeWorkbooksBytes({
      inputs: [input],
      includeSheetIndex: false,
    });
    expect(withoutIndex.result.warnings).toEqual([
      "1 source worksheet was hidden in the merged workbook.",
    ]);
    expect(withoutIndex.outputs[0]?.name).toBe("merged.xlsx");
    expect(
      XLSX.read(withoutIndex.outputs[0]!.bytes, { type: "array" }).SheetNames,
    ).toEqual(["Private"]);

    const reserved = await mergeWorkbooksBytes({
      inputs: [input],
      outputName: "aux.xlsx",
    });
    expect(reserved.outputs[0]?.name).toBe("_aux.xlsx");
  });

  it("replaces formulas with cached values when asked", async () => {
    const worksheet = XLSX.utils.aoa_to_sheet([
      ["Amount", "Tax", "Total"],
      [100, 5, { f: "A2+B2", t: "n", v: 105, z: "$#,##0.00" }],
    ]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Summary");
    const input = {
      name: "source.xlsx",
      bytes: new Uint8Array(
        XLSX.write(workbook, {
          bookType: "xlsx",
          cellStyles: true,
          compression: true,
          type: "array",
        }) as ArrayBuffer,
      ),
    };

    const formulas = await mergeWorkbooksBytes({ inputs: [input] });
    const values = await mergeWorkbooksBytes({
      inputs: [input],
      values: true,
    });
    const formulaCell = XLSX.read(formulas.outputs[0]!.bytes, {
      cellStyles: true,
      type: "array",
    }).Sheets.Summary?.C2;
    const valueCell = XLSX.read(values.outputs[0]!.bytes, {
      cellStyles: true,
      type: "array",
    }).Sheets.Summary?.C2;
    expect(formulaCell).toMatchObject({ f: "A2+B2", v: 105 });
    expect(valueCell).toMatchObject({ v: 105, z: "$#,##0.00" });
    expect(valueCell?.f).toBeUndefined();
  });

  it("produces byte-identical merges and cancels without output", async () => {
    const inputs = [
      {
        name: "north.xlsx",
        bytes: workbookBytes([["Summary", [["Region"], ["North"]]]]),
      },
    ];

    const first = await mergeWorkbooksBytes({ inputs });
    await new Promise((resolve) => setTimeout(resolve, CLOCK_TICK));
    const second = await mergeWorkbooksBytes({ inputs });
    expect(
      Buffer.compare(
        Buffer.from(first.outputs[0]!.bytes),
        Buffer.from(second.outputs[0]!.bytes),
      ),
    ).toBe(0);

    const controller = new AbortController();
    await expect(
      mergeWorkbooksBytes({
        inputs,
        signal: controller.signal,
        onProgress: () => controller.abort(),
      }),
    ).rejects.toMatchObject({ code: OPERATION_ABORTED });

    const aborted = new AbortController();
    aborted.abort();
    await expect(
      mergeWorkbooksBytes({ inputs, signal: aborted.signal }),
    ).rejects.toMatchObject({ code: OPERATION_ABORTED });
    await expect(mergeWorkbooksBytes({ inputs: [] })).rejects.toMatchObject({
      code: "XLSX_NO_INPUTS",
    });
    await expect(
      mergeWorkbooksBytes({
        inputs: [{ name: "corrupt.xlsx", bytes: corruptWorkbookBytes() }],
      }),
    ).rejects.toMatchObject({ code: "XLSX_READ_FAILED" });
  });
});

describe("byte-level worksheet records", () => {
  it("reads display text, skips empty rows, and validates the worksheet", async () => {
    const bytes = workbookBytes([
      [
        "Companies",
        [
          ["client", "amount"],
          ["A", 10],
          [null, null],
          ["B", 20],
        ],
      ],
    ]);

    const records = await readWorksheetRecordsBytes({
      name: "companies.xlsx",
      bytes,
    });
    expect(records.worksheet).toBe("Companies");
    expect(records.columns).toEqual(["client", "amount"]);
    expect(records.rows).toEqual([
      { amount: "10", client: "A" },
      { amount: "20", client: "B" },
    ]);
    expect(records.skippedEmptyRows).toBe(1);
    expect(records.sourceRows).toEqual([2, 4]);

    await expect(
      readWorksheetRecordsBytes(
        { name: "companies.xlsx", bytes },
        { worksheet: "Missing" },
      ),
    ).rejects.toMatchObject({ code: "XLSX_WORKSHEET_NOT_FOUND" });
  });
});

describe("byte entry point packaging", () => {
  it("keeps the built bytes entry free of node imports", async () => {
    const distDirectory = new URL("../dist/", import.meta.url);
    const visited = new Set<string>();
    const queue = ["bytes.js"];

    while (queue.length > 0) {
      const entry = queue.pop()!;
      if (visited.has(entry)) {
        continue;
      }
      visited.add(entry);
      const source = await readFile(
        fileURLToPath(new URL(entry, distDirectory)),
        "utf8",
      );
      expect(source, `${entry} must not import node builtins`).not.toMatch(
        /["']node:/u,
      );
      for (const match of source.matchAll(/from\s+["']\.\/([^"']+)["']/gu)) {
        queue.push(match[1]!);
      }
    }

    expect(visited.size).toBeGreaterThan(0);
  });
});
