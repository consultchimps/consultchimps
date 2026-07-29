import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { describe, expect, it } from "vitest";
import * as XLSX from "xlsx";

import { readWorksheetRecords } from "../src/index.js";

describe("readWorksheetRecords", () => {
  it("returns deterministic displayed text and skips only completely empty rows", async () => {
    const directory = await mkdtemp(
      path.join(tmpdir(), "consultchimps-xlsx-records-"),
    );

    try {
      const workbookPath = path.join(directory, "records.xlsx");
      const worksheet = XLSX.utils.aoa_to_sheet([
        ["name", "percentage", "date", "active", "empty", "formula"],
        [
          "Company A",
          0.125,
          new Date("2024-01-02T00:00:00.000Z"),
          true,
          null,
          null,
        ],
        [null, null, null, null, null, null],
        ["Company B", -0.021, null, false, null, null],
      ]);
      worksheet.B2!.z = "0.0%";
      worksheet.B4!.z = "0.0%";
      worksheet.C2!.z = "yyyy-mm-dd";
      worksheet.F2 = {
        f: "B2*2",
        t: "n",
        v: 0.25,
        z: "0.0%",
      };
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Companies");
      await writeFile(
        workbookPath,
        XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }),
      );

      await expect(
        readWorksheetRecords(workbookPath, {
          headerRow: 1,
          worksheet: "companies",
        }),
      ).resolves.toEqual({
        columns: ["name", "percentage", "date", "active", "empty", "formula"],
        rows: [
          {
            active: "TRUE",
            date: "2024-01-02",
            empty: "",
            formula: "25.0%",
            name: "Company A",
            percentage: "12.5%",
          },
          {
            active: "FALSE",
            date: "",
            empty: "",
            formula: "",
            name: "Company B",
            percentage: "-2.1%",
          },
        ],
        skippedEmptyRows: 1,
        sourceRows: [2, 4],
        worksheet: "Companies",
      });
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });
});
