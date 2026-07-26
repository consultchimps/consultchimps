import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { describe, expect, it } from "vitest";
import * as XLSX from "xlsx";

import { consolidateWorkbooks } from "../src/index.js";

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
    const directory = await mkdtemp(path.join(tmpdir(), "chimpcons-xlsx-"));

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

      const result = await consolidateWorkbooks([first, second], output);
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
});
