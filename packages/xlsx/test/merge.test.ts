import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { isConsultChimpsError } from "@consultchimps/core";
import { describe, expect, it } from "vitest";
import * as XLSX from "xlsx";

import { mergeWorkbooks, XLSX_ERRORS } from "../src/index.js";

async function createWorkbook(
  filePath: string,
  sheets: Array<{
    name: string;
    rows: Array<Array<number | string>>;
    visibility?: 0 | 1 | 2;
  }>,
): Promise<void> {
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
      Hidden: sheet.visibility ?? 0,
      name: sheet.name,
    })),
  };
  await writeFile(
    filePath,
    XLSX.write(workbook, {
      bookType: "xlsx",
      cellStyles: true,
      type: "buffer",
    }),
  );
}

describe("mergeWorkbooks", () => {
  it("replaces formulas with cached values without changing their number formats", async () => {
    const directory = await mkdtemp(
      path.join(tmpdir(), "consultchimps-merge-"),
    );

    try {
      const input = path.join(directory, "source.xlsx");
      const formulaOutput = path.join(directory, "formulas.xlsx");
      const valuesOutput = path.join(directory, "values.xlsx");
      const workbook = XLSX.utils.book_new();
      const worksheet = XLSX.utils.aoa_to_sheet([
        ["Amount", "Tax", "Total"],
        [100, 5, { f: "A2+B2", t: "n", v: 105, z: "$#,##0.00" }],
      ]);
      worksheet["!cols"] = [{ wch: 16 }, { wch: 12 }, { wch: 22 }];
      XLSX.utils.book_append_sheet(workbook, worksheet, "Summary");
      await writeFile(
        input,
        XLSX.write(workbook, {
          bookType: "xlsx",
          cellStyles: true,
          type: "buffer",
        }),
      );

      await mergeWorkbooks([input], formulaOutput);
      await mergeWorkbooks([input], valuesOutput, { values: true });

      const formulaWorkbook = XLSX.read(await readFile(formulaOutput), {
        cellStyles: true,
        type: "buffer",
      });
      const valuesWorkbook = XLSX.read(await readFile(valuesOutput), {
        cellStyles: true,
        type: "buffer",
      });
      expect(formulaWorkbook.Sheets.Summary?.C2).toMatchObject({
        f: "A2+B2",
        v: 105,
        z: "$#,##0.00",
      });
      expect(valuesWorkbook.Sheets.Summary?.C2).toMatchObject({
        v: 105,
        z: "$#,##0.00",
      });
      expect(valuesWorkbook.Sheets.Summary?.C2?.f).toBeUndefined();
      expect(valuesWorkbook.Sheets.Summary?.["!cols"]).toEqual(
        formulaWorkbook.Sheets.Summary?.["!cols"],
      );
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });

  it("copies worksheets in input order, resolves names, and records provenance", async () => {
    const directory = await mkdtemp(
      path.join(tmpdir(), "consultchimps-merge-"),
    );

    try {
      const first = path.join(directory, "north.xlsx");
      const second = path.join(directory, "south.xlsx");
      const output = path.join(directory, "outputs", "merged.xlsx");
      await createWorkbook(first, [
        { name: "Summary", rows: [["Region"], ["North"]] },
        {
          name: "Private",
          rows: [["Amount"], [100]],
          visibility: 2,
        },
      ]);
      await createWorkbook(second, [
        { name: "Summary", rows: [["Region"], ["South"]] },
        { name: "Sheet Index", rows: [["Existing source sheet"]] },
      ]);

      const result = await mergeWorkbooks([first, second], output);

      expect(result.metrics).toEqual({
        hiddenSheets: 1,
        inputFiles: 2,
        outputSheets: 4,
      });
      expect(result.warnings).toEqual([
        '1 source worksheet was hidden; see the visible "Sheet Index" worksheet.',
      ]);

      const workbook = XLSX.read(await readFile(output), { type: "buffer" });
      expect(workbook.SheetNames).toEqual([
        "Summary",
        "Private",
        "Summary (2)",
        "Sheet Index (2)",
        "Sheet Index",
      ]);
      expect(
        workbook.Workbook?.Sheets?.find((sheet) => sheet.name === "Private")
          ?.Hidden,
      ).toBe(2);
      expect(
        XLSX.utils.sheet_to_json(workbook.Sheets["Sheet Index"]!, {
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
        ["south.xlsx", "Sheet Index", "Sheet Index (2)", "Visible"],
      ]);
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });

  it("omits the index when requested and gives accurate hidden-sheet guidance", async () => {
    const directory = await mkdtemp(
      path.join(tmpdir(), "consultchimps-merge-"),
    );

    try {
      const input = path.join(directory, "source.xlsx");
      const output = path.join(directory, "merged.xlsx");
      await createWorkbook(input, [
        { name: "Private", rows: [["Value"], [1]], visibility: 1 },
      ]);

      const result = await mergeWorkbooks([input], output, {
        includeSheetIndex: false,
      });

      expect(result.warnings).toEqual([
        "1 source worksheet was hidden in the merged workbook.",
      ]);
      const workbook = XLSX.read(await readFile(output), { type: "buffer" });
      expect(workbook.SheetNames).toEqual(["Private"]);
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });

  it("validates inputs and destinations before writing", async () => {
    const directory = await mkdtemp(
      path.join(tmpdir(), "consultchimps-merge-"),
    );

    try {
      const input = path.join(directory, "source.xlsx");
      const output = path.join(directory, "existing.xlsx");
      await createWorkbook(input, [{ name: "Data", rows: [["Value"], [1]] }]);
      await createWorkbook(output, [
        { name: "Keep", rows: [["Do not overwrite"]] },
      ]);
      const existingBytes = await readFile(output);

      await expect(mergeWorkbooks([], output)).rejects.toMatchObject({
        code: XLSX_ERRORS.XLSX_NO_INPUTS,
      });
      await expect(mergeWorkbooks([input], input)).rejects.toMatchObject({
        code: "FILES_INPUT_OVERWRITE",
      });
      await expect(mergeWorkbooks([input], output)).rejects.toMatchObject({
        code: "FILES_OUTPUT_EXISTS",
      });
      expect(Buffer.compare(await readFile(output), existingBytes)).toBe(0);

      const invalid = path.join(directory, "missing.xlsx");
      const invalidOutput = path.join(directory, "new", "merged.xlsx");
      let thrown: unknown;
      try {
        await mergeWorkbooks([invalid], invalidOutput);
      } catch (error) {
        thrown = error;
      }
      expect(isConsultChimpsError(thrown)).toBe(true);
      expect(thrown).toMatchObject({ code: XLSX_ERRORS.XLSX_READ_FAILED });
      await expect(readFile(invalidOutput)).rejects.toThrow();
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });
});
