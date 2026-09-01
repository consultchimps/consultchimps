import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { isConsultChimpsError } from "@consultchimps/core";
import { describe, expect, it } from "vitest";
import * as XLSX from "xlsx";

import {
  consolidateWorkbooks,
  planConsolidateWorkbooks,
} from "../src/index.js";

async function createWorkbook(
  filePath: string,
  sheetName: string,
  rows: Array<Array<XLSX.CellObject | Date | string | number>>,
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

async function writeMapping(
  filePath: string,
  document: unknown,
): Promise<string> {
  await writeFile(filePath, JSON.stringify(document, null, 2), "utf8");
  return filePath;
}

function readGrid(bytes: Buffer, sheetName: string): unknown[][] {
  const workbook = XLSX.read(bytes, { type: "buffer" });
  const worksheet = workbook.Sheets[sheetName];
  expect(worksheet).toBeDefined();
  return XLSX.utils.sheet_to_json(worksheet!, {
    defval: null,
    header: 1,
    raw: true,
  });
}

describe("consolidateWorkbooks with a column mapping", () => {
  it("folds aliases, coerces values, and warns about unmapped columns", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-xlsx-"));

    try {
      const north = path.join(directory, "north.xlsx");
      const south = path.join(directory, "south.xlsx");
      const output = path.join(directory, "combined.xlsx");
      await createWorkbook(north, "Cases", [
        ["Case ID", "Total", "Run Date", "Region"],
        ["R-1", "1.234,50", "09/03/2024", "north"],
      ]);
      await createWorkbook(south, "Cases", [
        ["reference", "amount", "Region"],
        ["R-2", "7,25", "south"],
      ]);
      const mappingFile = await writeMapping(
        path.join(directory, "mapping.json"),
        {
          version: 1,
          columns: [
            { name: "Case_ID", aliases: ["Reference", "Case Number"] },
            {
              name: "Amount",
              aliases: ["Total"],
              coercion: {
                type: "number",
                decimalSeparator: ",",
                thousandsSeparator: ".",
              },
            },
            {
              name: "Opened_On",
              aliases: ["Run Date"],
              coercion: { type: "date", format: "DD/MM/YYYY" },
            },
          ],
          constants: { Dataset: "quarterly" },
        },
      );

      const result = await consolidateWorkbooks({
        inputs: [north, south],
        output,
        addSourceColumns: false,
        mappingFile,
      });

      // "Case ID" and "reference" both reach Case_ID by normalized key, and
      // "amount" reaches Amount without being listed as an alias, because a
      // canonical column always matches its own name.
      expect(result.metrics).toMatchObject({
        outputColumns: 5,
        outputRows: 2,
        suggestedColumns: 0,
        unmappedColumns: 1,
      });
      expect(result.warnings).toEqual([
        '1 column did not match the column mapping and kept its own name: "Region". Add it to the mapping if it belongs in a canonical column.',
      ]);
      expect(result.suggestion).toBeUndefined();

      expect(readGrid(await readFile(output), "Consolidated")).toEqual([
        ["Case_ID", "Amount", "Opened_On", "Region", "Dataset"],
        ["R-1", 1234.5, "2024-03-09", "north", "quarterly"],
        ["R-2", 7.25, null, "south", "quarterly"],
      ]);
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });

  it("matches by normalized key whether or not headers are normalized", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-xlsx-"));

    try {
      const north = path.join(directory, "north.xlsx");
      const south = path.join(directory, "south.xlsx");
      await createWorkbook(north, "Cases", [
        ["Case ID", "Failed Checks"],
        ["R-1", 5],
      ]);
      await createWorkbook(south, "Cases", [
        ["Case_ID", "Failed_Checks"],
        ["R-2", 7],
      ]);
      const mappingFile = await writeMapping(path.join(directory, "map.json"), {
        version: 1,
        columns: [{ name: "Case_ID", aliases: [] }],
      });

      // The mapping folds both spellings of the case column either way; only
      // the columns it did not claim answer to normalizeHeaders.
      const exact = await consolidateWorkbooks({
        inputs: [north, south],
        output: path.join(directory, "exact.xlsx"),
        addSourceColumns: false,
        mappingFile,
      });
      expect(exact.metrics.outputColumns).toBe(3);
      expect(exact.metrics.unmappedColumns).toBe(2);
      expect(
        readGrid(
          await readFile(path.join(directory, "exact.xlsx")),
          "Consolidated",
        )[0],
      ).toEqual(["Case_ID", "Failed Checks", "Failed_Checks"]);

      const normalized = await consolidateWorkbooks({
        inputs: [north, south],
        output: path.join(directory, "normalized.xlsx"),
        addSourceColumns: false,
        mappingFile,
        normalizeHeaders: true,
      });
      expect(normalized.metrics.outputColumns).toBe(2);
      expect(normalized.metrics.unmappedColumns).toBe(2);
      expect(
        readGrid(
          await readFile(path.join(directory, "normalized.xlsx")),
          "Consolidated",
        )[0],
      ).toEqual(["Case_ID", "Failed Checks"]);
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });

  it("refuses two columns of one worksheet folding into one canonical column", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-xlsx-"));

    try {
      const input = path.join(directory, "north.xlsx");
      const output = path.join(directory, "combined.xlsx");
      await createWorkbook(input, "Cases", [
        ["Case ID", "Reference"],
        ["R-1", "R-1-alt"],
      ]);
      const mappingFile = await writeMapping(path.join(directory, "map.json"), {
        version: 1,
        columns: [{ name: "Case_ID", aliases: ["Case ID", "Reference"] }],
      });

      await expect(
        consolidateWorkbooks({ inputs: [input], output, mappingFile }),
      ).rejects.toMatchObject({
        code: "TABLE_MAPPING_COLUMN_COLLISION",
        details: { canonicalColumn: "Case_ID", sheet: "Cases" },
      });
      // The refusal happens before the writer runs, so no output exists.
      await expect(readFile(output)).rejects.toMatchObject({ code: "ENOENT" });
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });

  it("refuses a date coercion over a numeric cell instead of reading it as a serial", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-xlsx-"));

    try {
      const input = path.join(directory, "north.xlsx");
      const output = path.join(directory, "combined.xlsx");
      // 45360 is the 1900-system serial for 9 March 2024 and also a perfectly
      // ordinary number. Which one it is belongs to the workbook, not to the
      // cell, so the run stops rather than guessing an epoch.
      await createWorkbook(input, "Cases", [["Run Date"], [45360]]);
      const mappingFile = await writeMapping(path.join(directory, "map.json"), {
        version: 1,
        columns: [
          {
            name: "Opened_On",
            aliases: ["Run Date"],
            coercion: { type: "date", format: "DD/MM/YYYY" },
          },
        ],
      });

      const failure: unknown = await consolidateWorkbooks({
        inputs: [input],
        output,
        mappingFile,
      }).then(
        () => undefined,
        (error: unknown) => error,
      );
      if (!isConsultChimpsError(failure)) {
        throw new Error("Expected a ConsultChimps error.");
      }
      expect(failure.code).toBe("XLSX_MAPPING_DATE_NOT_TEXT");
      expect(failure.message).toContain('column "Run Date"');
      expect(failure.message).toContain("holds the number 45360");
      expect(failure.details).toMatchObject({
        canonicalColumn: "Opened_On",
        column: "Run Date",
        format: "DD/MM/YYYY",
        row: 2,
        sheet: "Cases",
        valueType: "number",
      });
      await expect(readFile(output)).rejects.toMatchObject({ code: "ENOENT" });
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });

  it("refuses a date coercion over a cell the workbook already stores as a date", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-xlsx-"));

    try {
      const input = path.join(directory, "north.xlsx");
      const output = path.join(directory, "combined.xlsx");
      await createWorkbook(input, "Cases", [
        ["Run Date"],
        [new Date(Date.UTC(2024, 2, 9, 12))],
      ]);
      const mappingFile = await writeMapping(path.join(directory, "map.json"), {
        version: 1,
        columns: [
          {
            name: "Opened_On",
            aliases: ["Run Date"],
            coercion: { type: "date", format: "DD/MM/YYYY" },
          },
        ],
      });

      await expect(
        consolidateWorkbooks({ inputs: [input], output, mappingFile }),
      ).rejects.toMatchObject({
        code: "XLSX_MAPPING_DATE_NOT_TEXT",
        details: { column: "Run Date", valueType: "workbook-date" },
      });
      await expect(readFile(output)).rejects.toMatchObject({ code: "ENOENT" });
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });

  it("carries text dates through the column that refuses a serial", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-xlsx-"));

    try {
      const input = path.join(directory, "north.xlsx");
      const output = path.join(directory, "combined.xlsx");
      await createWorkbook(input, "Cases", [
        ["Run Date", "Case ID"],
        [{ t: "s", v: "09/03/2024" }, "R-1"],
        [{ t: "s", v: "" }, "R-2"],
      ]);
      const mappingFile = await writeMapping(path.join(directory, "map.json"), {
        version: 1,
        columns: [
          {
            name: "Opened_On",
            aliases: ["Run Date"],
            coercion: { type: "date", format: "DD/MM/YYYY" },
          },
        ],
      });

      const result = await consolidateWorkbooks({
        inputs: [input],
        output,
        addSourceColumns: false,
        mappingFile,
      });
      expect(result.metrics.outputRows).toBe(2);
      // A blank cell stays blank; text in the declared format becomes ISO.
      expect(readGrid(await readFile(output), "Consolidated")).toEqual([
        ["Opened_On", "Case ID"],
        ["2024-03-09", "R-1"],
        [null, "R-2"],
      ]);
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });

  it("reports an unreadable, unparseable, or ambiguous mapping before reading a workbook", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-xlsx-"));

    try {
      const output = path.join(directory, "combined.xlsx");
      // The input does not exist: reaching a mapping failure proves the
      // mapping is settled before any workbook is opened.
      const missingInput = path.join(directory, "absent.xlsx");

      await expect(
        consolidateWorkbooks({
          inputs: [missingInput],
          output,
          mappingFile: path.join(directory, "absent.json"),
        }),
      ).rejects.toMatchObject({ code: "XLSX_MAPPING_FILE_UNREADABLE" });

      const notJson = path.join(directory, "not-json.json");
      await writeFile(notJson, "{ this is not JSON", "utf8");
      await expect(
        consolidateWorkbooks({
          inputs: [missingInput],
          output,
          mappingFile: notJson,
        }),
      ).rejects.toMatchObject({ code: "XLSX_MAPPING_FILE_INVALID" });

      const ambiguous = await writeMapping(
        path.join(directory, "ambiguous.json"),
        {
          version: 1,
          columns: [
            { name: "Case_ID", aliases: ["Reference"] },
            { name: "Other", aliases: ["reference"] },
          ],
        },
      );
      await expect(
        consolidateWorkbooks({
          inputs: [missingInput],
          output,
          mappingFile: ambiguous,
        }),
      ).rejects.toMatchObject({
        code: "TABLE_MAPPING_INVALID",
        details: { problem: "duplicate_alias" },
      });
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });

  it("writes a drafted mapping beside the consolidation and refuses to replace one", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-xlsx-"));

    try {
      const north = path.join(directory, "north.xlsx");
      const south = path.join(directory, "south.xlsx");
      const output = path.join(directory, "combined.xlsx");
      const draft = path.join(directory, "drafts", "mapping.json");
      await createWorkbook(north, "Cases", [
        ["Failed Checks", "Case ID"],
        [5, "R-1"],
      ]);
      await createWorkbook(south, "Cases", [
        ["Failed_Checks", "Case ID"],
        [7, "R-2"],
      ]);

      const result = await consolidateWorkbooks({
        inputs: [north, south],
        output,
        suggestMappingOutput: draft,
      });

      // Only the group that actually disagrees about spelling is proposed:
      // "Case ID" is spelled the same way everywhere and needs no entry.
      expect(result.metrics).toMatchObject({
        suggestedColumns: 1,
        unmappedColumns: 0,
      });
      expect(result.suggestion?.mapping).toEqual({
        version: 1,
        columns: [{ name: "Failed Checks", aliases: [] }],
      });
      expect(result.suggestion?.groups[0]?.spellings).toEqual([
        "Failed Checks",
        "Failed_Checks",
      ]);
      expect(result.artifacts).toEqual([
        {
          kind: "file",
          mediaType:
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          path: output,
        },
        { kind: "file", mediaType: "application/json", path: draft },
      ]);
      expect(JSON.parse(await readFile(draft, "utf8"))).toEqual({
        version: 1,
        columns: [{ name: "Failed Checks", aliases: [] }],
      });
      // The consolidation still runs: the draft is evidence for a second run,
      // not a mode that replaces this one.
      expect(result.metrics.outputRows).toBe(2);

      const again = path.join(directory, "again.xlsx");
      await expect(
        consolidateWorkbooks({
          inputs: [north, south],
          output: again,
          suggestMappingOutput: draft,
        }),
      ).rejects.toMatchObject({ code: "FILES_OUTPUT_EXISTS" });
      // The destination is checked before the workbooks are read, so the
      // refused run left nothing behind.
      await expect(readFile(again)).rejects.toMatchObject({ code: "ENOENT" });
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });

  it("refuses to apply and to draft a mapping in one run", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-xlsx-"));

    try {
      const input = path.join(directory, "north.xlsx");
      await createWorkbook(input, "Cases", [["Case ID"], ["R-1"]]);
      const mappingFile = await writeMapping(path.join(directory, "map.json"), {
        version: 1,
        columns: [{ name: "Case_ID", aliases: ["Case ID"] }],
      });

      await expect(
        consolidateWorkbooks({
          inputs: [input],
          output: path.join(directory, "combined.xlsx"),
          mappingFile,
          suggestMappingOutput: path.join(directory, "draft.json"),
        }),
      ).rejects.toMatchObject({
        code: "XLSX_MAPPING_SUGGEST_CONFLICT",
        details: { problem: "mapping_and_suggestion" },
      });

      await expect(
        consolidateWorkbooks({
          inputs: [input],
          output: path.join(directory, "combined.xlsx"),
          suggestMappingOutput: path.join(directory, "combined.xlsx"),
        }),
      ).rejects.toMatchObject({
        code: "XLSX_MAPPING_SUGGEST_CONFLICT",
        details: { problem: "suggestion_shares_output" },
      });
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });

  it("refuses destinations that collide without being spelled alike", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-xlsx-"));

    try {
      const input = path.join(directory, "north.xlsx");
      await createWorkbook(input, "Cases", [["Case ID"], ["R-1"]]);
      const output = path.join(directory, "combined.xlsx");

      // One path cannot be both a file and a directory, so a draft nested
      // under the workbook destination is refused before either is written.
      await expect(
        consolidateWorkbooks({
          inputs: [input],
          output,
          overwrite: true,
          suggestMappingOutput: path.join(output, "mapping.json"),
        }),
      ).rejects.toMatchObject({
        code: "XLSX_MAPPING_SUGGEST_CONFLICT",
        details: { problem: "suggestion_shares_output" },
      });
      await expect(readFile(output)).rejects.toMatchObject({ code: "ENOENT" });

      // On a case-folding filesystem these two names are one file. The check
      // asks the platform's question, so the refusal happens there and the
      // two genuinely distinct files are allowed through on Linux.
      const caseVariant = consolidateWorkbooks({
        inputs: [input],
        output,
        overwrite: true,
        suggestMappingOutput: path.join(directory, "Combined.xlsx"),
      });
      if (process.platform === "win32" || process.platform === "darwin") {
        await expect(caseVariant).rejects.toMatchObject({
          code: "XLSX_MAPPING_SUGGEST_CONFLICT",
          details: { problem: "suggestion_shares_output" },
        });
        await expect(readFile(output)).rejects.toMatchObject({
          code: "ENOENT",
        });
      } else {
        expect((await caseVariant).artifacts).toHaveLength(2);
      }
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });

  it("refuses an output aimed at the mapping file it was told to read", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-xlsx-"));

    try {
      const input = path.join(directory, "north.xlsx");
      await createWorkbook(input, "Cases", [["Case ID"], ["R-1"]]);
      const mappingFile = await writeMapping(path.join(directory, "map.json"), {
        version: 1,
        columns: [{ name: "Case_ID", aliases: ["Case ID"] }],
      });

      // The mapping is an input of this run, so overwrite must not reach it.
      await expect(
        consolidateWorkbooks({
          inputs: [input],
          output: mappingFile,
          mappingFile,
          overwrite: true,
        }),
      ).rejects.toMatchObject({ code: "FILES_INPUT_OVERWRITE" });
      expect(JSON.parse(await readFile(mappingFile, "utf8"))).toMatchObject({
        version: 1,
      });
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });

  it("plans the drafted mapping as a second output and checks the mapping first", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-xlsx-"));

    try {
      const input = path.join(directory, "north.xlsx");
      await createWorkbook(input, "Cases", [["Case ID"], ["R-1"]]);
      const draft = path.join(directory, "draft.json");

      const plan = await planConsolidateWorkbooks({
        inputs: [input],
        output: path.join(directory, "combined.xlsx"),
        suggestMappingOutput: draft,
      });
      expect(plan.outputs).toHaveLength(2);
      expect(plan.outputs[1]).toEqual({
        kind: "file",
        mediaType: "application/json",
        path: draft,
        exists: false,
      });
      expect(plan.metrics.outputFiles).toBe(2);

      await expect(
        planConsolidateWorkbooks({
          inputs: [input],
          output: path.join(directory, "combined.xlsx"),
          mappingFile: path.join(directory, "absent.json"),
        }),
      ).rejects.toMatchObject({ code: "XLSX_MAPPING_FILE_UNREADABLE" });

      // A plan lists every file the run reads, so the mapping joins the
      // workbooks without changing the workbook count.
      const mappingFile = await writeMapping(path.join(directory, "map.json"), {
        version: 1,
        columns: [{ name: "Case_ID", aliases: ["Case ID"] }],
      });
      const mapped = await planConsolidateWorkbooks({
        inputs: [input],
        output: path.join(directory, "combined.xlsx"),
        mappingFile,
      });
      expect(mapped.inputs).toEqual([input, mappingFile]);
      expect(mapped.metrics.inputFiles).toBe(1);
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });

  it("refuses a draft whose folder is already a file, before writing the workbook", async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-xlsx-"));

    try {
      const input = path.join(directory, "north.xlsx");
      await createWorkbook(input, "Cases", [["Case ID"], ["R-1"]]);
      const output = path.join(directory, "combined.xlsx");
      // A plain file already stands where the draft's folder would go. Only
      // trying to create the folder reveals that, so it is tried up front.
      const blocker = path.join(directory, "drafts");
      await writeFile(blocker, "not a folder", "utf8");

      await expect(
        consolidateWorkbooks({
          inputs: [input],
          output,
          suggestMappingOutput: path.join(blocker, "mapping.json"),
        }),
      ).rejects.toThrow();
      await expect(readFile(output)).rejects.toMatchObject({ code: "ENOENT" });
      expect(await readFile(blocker, "utf8")).toBe("not a folder");
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });
});
