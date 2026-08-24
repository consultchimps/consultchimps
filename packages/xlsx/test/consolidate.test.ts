import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import {
  isConsultChimpsError,
  OPERATION_ABORTED,
  type OperationProgress,
} from "@consultchimps/core";
import { describe, expect, it } from "vitest";
import * as XLSX from "xlsx";

import {
  consolidateWorkbooks,
  planConsolidateWorkbooks,
} from "../src/index.js";

async function createWorkbook(
  filePath: string,
  sheetName: string,
  rows: Array<Array<XLSX.CellObject | string | number>>,
): Promise<void> {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet(rows),
    sheetName,
  );
  await writeFile(
    filePath,
    XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }),
  );
}

describe("consolidateWorkbooks", () => {
  it("combines worksheets with different column order and provenance", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-xlsx-"));

    try {
      const first = path.join(directory, "north.xlsx");
      const second = path.join(directory, "south.xlsx");
      const output = path.join(directory, "consolidated.xlsx");
      await createWorkbook(first, "North", [
        ["Client", "Amount"],
        ["A", { f: "5+5", t: "n", v: 10 }],
      ]);
      await createWorkbook(second, "South", [
        ["Amount", "Status", "client"],
        [20, "Open", "B"],
      ]);

      const result = await consolidateWorkbooks({
        inputs: [first, second],
        output: output,
        values: true,
      });
      expect(result.metrics).toMatchObject({
        inputFiles: 2,
        inputTables: 2,
        outputRows: 2,
      });

      const workbook = XLSX.read(await readFile(output), { type: "buffer" });
      const worksheet = workbook.Sheets.Consolidated;
      expect(worksheet).toBeDefined();
      expect(
        XLSX.utils.sheet_to_json(worksheet!, {
          defval: null,
          header: 1,
          raw: true,
        }),
      ).toEqual([
        [
          "Client",
          "Amount",
          "Status",
          "_source_file",
          "_source_sheet",
          "_source_row",
        ],
        ["A", 10, null, "north.xlsx", "North", 2],
        ["B", 20, "Open", "south.xlsx", "South", 2],
      ]);
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });

  it("reports deterministic progress and honours cancellation", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-xlsx-"));

    try {
      const first = path.join(directory, "north.xlsx");
      const second = path.join(directory, "south.xlsx");
      const output = path.join(directory, "consolidated.xlsx");
      await createWorkbook(first, "North", [
        ["Client", "Amount"],
        ["A", 10],
      ]);
      await createWorkbook(second, "South", [
        ["Client", "Amount"],
        ["B", 20],
      ]);

      const events: OperationProgress[] = [];
      await consolidateWorkbooks({
        inputs: [first, second],
        output,
        onProgress: (progress) => events.push(progress),
      });
      expect(events.map((event) => [event.stage, event.completed])).toEqual([
        ["reading-workbooks", 1],
        ["reading-workbooks", 2],
        ["writing-output", 1],
      ]);

      const controller = new AbortController();
      controller.abort();
      let thrown: unknown;
      try {
        await consolidateWorkbooks({
          inputs: [first, second],
          output: path.join(directory, "cancelled.xlsx"),
          signal: controller.signal,
        });
      } catch (error) {
        thrown = error;
      }
      expect(isConsultChimpsError(thrown)).toBe(true);
      expect((thrown as { code: string }).code).toBe(OPERATION_ABORTED);
      await expect(
        readFile(path.join(directory, "cancelled.xlsx")),
      ).rejects.toThrow();
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });

  it("plans the consolidation without writing and flags collisions", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-xlsx-"));

    try {
      const first = path.join(directory, "north.xlsx");
      const output = path.join(directory, "consolidated.xlsx");
      await createWorkbook(first, "North", [
        ["Client", "Amount"],
        ["A", 10],
      ]);

      const plan = await planConsolidateWorkbooks({
        inputs: [first],
        output,
      });
      expect(plan.operation).toBe("sheets.consolidate");
      expect(plan.outputs).toEqual([
        {
          kind: "file",
          mediaType:
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          path: path.resolve(output),
          exists: false,
        },
      ]);
      expect(plan.warnings).toEqual([]);
      await expect(readFile(output)).rejects.toThrow();

      await createWorkbook(output, "Existing", [["Header"]]);
      const collidingPlan = await planConsolidateWorkbooks({
        inputs: [first],
        output,
      });
      expect(collidingPlan.outputs[0]?.exists).toBe(true);
      expect(collidingPlan.warnings).toHaveLength(1);

      await expect(
        planConsolidateWorkbooks({
          inputs: [path.join(directory, "missing.xlsx")],
          output: path.join(directory, "other.xlsx"),
        }),
      ).rejects.toThrowError(/Workbook not found/);
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });

  it("normalizes header variants across differently exported workbooks", async () => {
    // Simulates one schema exported by different systems: the same headers
    // written with spaces, underscores, colon spacing, case drift, trailing
    // spaces, extra columns, and a hidden summary sheet.
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-xlsx-"));

    const createLogWorkbook = async (
      filePath: string,
      sheets: Array<{
        hidden?: boolean;
        name: string;
        rows: Array<Array<string | number>>;
      }>,
    ): Promise<void> => {
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
          Hidden: sheet.hidden === true ? 1 : 0,
        })),
      };
      await writeFile(
        filePath,
        XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }),
      );
    };

    try {
      const alpha = path.join(directory, "north.xlsx");
      const bravo = path.join(directory, "south.xlsx");
      const charlie = path.join(directory, "east.xlsx");
      const output = path.join(directory, "consolidated.xlsx");

      await createLogWorkbook(alpha, [
        {
          name: "Review Log",
          rows: [
            [
              "Case_ID",
              "Failed Checks",
              "Total Checks",
              "Reviewer: Lead Contact",
            ],
            ["R-1", 5, 100, "Reviewer A"],
          ],
        },
        {
          hidden: true,
          name: "Summary",
          rows: [
            ["Category", "Count"],
            ["Complete", 3],
          ],
        },
      ]);
      await createLogWorkbook(bravo, [
        {
          name: "vF",
          rows: [
            [
              "S.No.",
              "Case_ID",
              "Failed_Checks",
              "Total_Checks",
              "Reviewer_Lead_Contact",
            ],
            [1, "R-2", 7, 200, "Reviewer B"],
          ],
        },
      ]);
      await createLogWorkbook(charlie, [
        {
          name: "Sheet1",
          rows: [
            [
              "Case_ID",
              "Failed Checks ",
              "Total  Checks",
              "Reviewer:Lead Contact",
            ],
            ["R-3", 9, 300, "Reviewer C"],
          ],
        },
        {
          name: "Lookup",
          rows: [["Case_ID"], ["R-4"]],
        },
      ]);

      const result = await consolidateWorkbooks({
        inputs: [alpha, bravo, charlie],
        output,
        normalizeHeaders: true,
      });
      expect(result.metrics).toMatchObject({
        inputFiles: 3,
        inputTables: 4,
        outputColumns: 8,
        outputRows: 4,
      });

      const workbook = XLSX.read(await readFile(output), { type: "buffer" });
      expect(
        XLSX.utils.sheet_to_json(workbook.Sheets.Consolidated!, {
          defval: null,
          header: 1,
          raw: true,
        }),
      ).toEqual([
        [
          "Case_ID",
          "Failed Checks",
          "Total Checks",
          "Reviewer: Lead Contact",
          "S.No.",
          "_source_file",
          "_source_sheet",
          "_source_row",
        ],
        ["R-1", 5, 100, "Reviewer A", null, "north.xlsx", "Review Log", 2],
        ["R-2", 7, 200, "Reviewer B", 1, "south.xlsx", "vF", 2],
        ["R-3", 9, 300, "Reviewer C", null, "east.xlsx", "Sheet1", 2],
        ["R-4", null, null, null, null, "east.xlsx", "Lookup", 2],
      ]);

      // Without the option the variants stay separate - the behaviour that
      // turns one shared schema into a doubled-up column list.
      const exact = await consolidateWorkbooks({
        inputs: [alpha, bravo, charlie],
        output: path.join(directory, "exact.xlsx"),
      });
      expect(exact.metrics.outputColumns).toBe(13);
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });

  it("produces byte-identical output for identical inputs", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-xlsx-"));

    try {
      const input = path.join(directory, "north.xlsx");
      await createWorkbook(input, "North", [
        ["Client", "Amount"],
        ["A", 10],
      ]);

      const first = path.join(directory, "first.xlsx");
      const second = path.join(directory, "second.xlsx");
      await consolidateWorkbooks({ inputs: [input], output: first });
      await new Promise((resolve) => setTimeout(resolve, 1100));
      await consolidateWorkbooks({ inputs: [input], output: second });

      expect(
        Buffer.compare(await readFile(first), await readFile(second)),
      ).toBe(0);
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });
});
