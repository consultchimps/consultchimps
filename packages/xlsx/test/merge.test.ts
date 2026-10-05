import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import {
  isConsultChimpsError,
  OPERATION_ABORTED,
  type OperationProgress,
} from "@consultchimps/core";
import { describe, expect, it } from "vitest";

import { mergeWorkbooks, XLSX_ERRORS } from "../src/index.js";
import {
  readPackagePart,
  styleNumberFormatCode,
  worksheetCellFormula,
  worksheetCellStyle,
  worksheetCellValue,
} from "./corpus/fixtures.js";
import {
  sheetNamesOf,
  sheetRows,
  sheetVisibilities,
  sheetXml,
} from "./support/read-workbook.js";
import { buildWorkbookFixture } from "./support/workbook-fixture.js";

const VISIBILITY = [undefined, "hidden", "veryHidden"] as const;

/** One cell's formula, cached value and number format, from the raw parts. */
async function cellFacts(
  bytes: Uint8Array,
  sheetName: string,
  reference: string,
): Promise<Record<string, string | undefined>> {
  const xml = await sheetXml(bytes, sheetName);
  const styles = await readPackagePart(bytes, "xl/styles.xml");
  return {
    formula: worksheetCellFormula(xml, reference),
    value: worksheetCellValue(xml, reference),
    format: styleNumberFormatCode(styles, worksheetCellStyle(xml, reference)),
    cols: /<cols>[\s\S]*?<\/cols>/u.exec(xml)?.[0],
  };
}

async function createWorkbook(
  filePath: string,
  sheets: Array<{
    name: string;
    rows: Array<Array<number | string>>;
    visibility?: 0 | 1 | 2;
  }>,
): Promise<void> {
  await writeFile(
    filePath,
    await buildWorkbookFixture({
      sheets: sheets.map((sheet) => {
        const state = VISIBILITY[sheet.visibility ?? 0];
        return {
          name: sheet.name,
          rows: sheet.rows,
          ...(state ? { state } : {}),
        };
      }),
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
      await writeFile(
        input,
        await buildWorkbookFixture({
          sheets: [
            {
              name: "Summary",
              rows: [
                ["Amount", "Tax", "Total"],
                [100, 5, { formula: "A2+B2", value: 105, format: "$#,##0.00" }],
              ],
              widths: [16, 12, 22],
            },
          ],
        }),
      );

      await mergeWorkbooks([input], formulaOutput);
      await mergeWorkbooks([input], valuesOutput, { values: true });

      const formulaCell = await cellFacts(
        await readFile(formulaOutput),
        "Summary",
        "C2",
      );
      const valuesCell = await cellFacts(
        await readFile(valuesOutput),
        "Summary",
        "C2",
      );
      expect(formulaCell).toMatchObject({
        formula: "A2+B2",
        value: "105",
        format: "$#,##0.00",
      });
      expect(valuesCell).toMatchObject({ value: "105", format: "$#,##0.00" });
      expect(valuesCell.formula).toBeUndefined();
      expect(formulaCell.cols).toContain('width="16"');
      expect(valuesCell.cols).toEqual(formulaCell.cols);
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

      const bytes = await readFile(output);
      expect(await sheetNamesOf(bytes)).toEqual([
        "Summary",
        "Private",
        "Summary (2)",
        "Sheet Index (2)",
        "Sheet Index",
      ]);
      expect((await sheetVisibilities(bytes)).Private).toBe("veryHidden");
      expect(await sheetRows(bytes, "Sheet Index")).toEqual([
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

  it("reports deterministic progress and honours cancellation", async () => {
    const directory = await mkdtemp(
      path.join(tmpdir(), "consultchimps-merge-"),
    );

    try {
      const first = path.join(directory, "north.xlsx");
      const second = path.join(directory, "south.xlsx");
      const output = path.join(directory, "merged.xlsx");
      await createWorkbook(first, [
        { name: "Summary", rows: [["Region"], ["North"]] },
      ]);
      await createWorkbook(second, [
        { name: "Summary", rows: [["Region"], ["South"]] },
      ]);

      const events: OperationProgress[] = [];
      await mergeWorkbooks([first, second], output, {
        onProgress: (progress) => events.push(progress),
      });
      expect(events).toEqual([
        {
          operation: "sheets.merge",
          stage: "merging-inputs",
          completed: 1,
          total: 2,
          detail: "north.xlsx",
        },
        {
          operation: "sheets.merge",
          stage: "merging-inputs",
          completed: 2,
          total: 2,
          detail: "south.xlsx",
        },
        {
          operation: "sheets.merge",
          stage: "writing-output",
          completed: 1,
          total: 1,
          detail: "merged.xlsx",
        },
      ]);

      const controller = new AbortController();
      controller.abort();
      let thrown: unknown;
      try {
        await mergeWorkbooks(
          [first, second],
          path.join(directory, "cancelled.xlsx"),
          { signal: controller.signal },
        );
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
      expect(await sheetNamesOf(await readFile(output))).toEqual(["Private"]);
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
