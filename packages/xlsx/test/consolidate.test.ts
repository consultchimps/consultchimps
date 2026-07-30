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
  rows: Array<Array<string | number>>,
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
        ["A", 10],
      ]);
      await createWorkbook(second, "South", [
        ["Amount", "Status", "client"],
        [20, "Open", "B"],
      ]);

      const result = await consolidateWorkbooks({
        inputs: [first, second],
        output: output,
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
});
