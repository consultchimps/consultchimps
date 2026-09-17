import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { ConsultChimpsError } from "@consultchimps/core";
import { PipelineBudget, validateExportOptions } from "../src/budget.js";
import { exportPbiTables, readPbiTables } from "../src/pipeline.js";
import { FETCHED, FIXTURES } from "./oracle.js";

/**
 * The export path end to end on the smallest committed fixture: the two
 * artifacts, their names, media types and order, the six metrics, the manifest
 * shape, and byte determinism across two runs in one process.
 */

const fixture = new Uint8Array(
  readFileSync(path.join(FIXTURES, "a-2018-fuzzy.pbix")),
);

async function sheetNames(bytes: Uint8Array): Promise<string[]> {
  const zip = await JSZip.loadAsync(bytes);
  const workbook = await zip.file("xl/workbook.xml")!.async("string");
  return [...workbook.matchAll(/<sheet name="([^"]*)"/g)].map(
    (match) => match[1]!,
  );
}

describe("exportPbiTables", () => {
  it("produces a workbook and a manifest, in that order, with the six metrics", async () => {
    const outcome = await exportPbiTables(fixture);
    expect(outcome.outputs.map((output) => output.name)).toEqual([
      "power-bi-tables.xlsx",
      "power-bi-tables.manifest.json",
    ]);
    expect(outcome.outputs.map((output) => output.mediaType)).toEqual([
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "application/json",
    ]);
    expect(outcome.result.artifacts).toEqual([
      {
        kind: "file",
        path: "power-bi-tables.xlsx",
        mediaType:
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      },
      {
        kind: "file",
        path: "power-bi-tables.manifest.json",
        mediaType: "application/json",
      },
    ]);
    expect(Object.keys(outcome.result.metrics)).toEqual([
      "inputFiles",
      "outputFiles",
      "exportedTables",
      "exportedColumns",
      "exportedRows",
      "outputWorksheets",
    ]);
    // Two visible tables; the hidden date template is skipped by default.
    expect(outcome.result.metrics).toEqual({
      inputFiles: 1,
      outputFiles: 2,
      exportedTables: 2,
      exportedColumns: 9,
      exportedRows: 22,
      outputWorksheets: 2,
    });
    expect(await sheetNames(outcome.outputs[0]!.bytes)).toEqual([
      "People",
      "Sales",
    ]);
  }, 120_000);

  it("names the include-hidden option in the hidden-table warning", async () => {
    const outcome = await exportPbiTables(fixture);
    expect(
      outcome.result.warnings.some(
        (warning) =>
          warning.includes("hidden") && warning.includes("includeHiddenTables"),
      ),
    ).toBe(true);
  }, 120_000);

  it("includes hidden tables when asked", async () => {
    const outcome = await exportPbiTables(fixture, {
      includeHiddenTables: true,
    });
    expect(outcome.result.metrics.exportedTables).toBe(3);
    expect(outcome.result.metrics.exportedRows).toBe(23);
  }, 120_000);

  it("writes a manifest with a fixed schema version and no timestamps", async () => {
    const outcome = await exportPbiTables(fixture);
    const manifest = JSON.parse(
      new TextDecoder().decode(outcome.outputs[1]!.bytes),
    ) as Record<string, unknown>;
    expect(Object.keys(manifest)).toEqual([
      "schemaVersion",
      "tables",
      "excludedTables",
      "unverifiedPaths",
    ]);
    expect(manifest.schemaVersion).toBe(1);
    const text = new TextDecoder().decode(outcome.outputs[1]!.bytes);
    expect(text).not.toMatch(/\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/);
    const tables = manifest.tables as { name: string; parts: unknown[] }[];
    expect(tables.map((table) => table.name)).toEqual(["People", "Sales"]);
    expect(tables[0]!.parts).toEqual([
      { sheetName: "People", start: 0, end: 13 },
    ]);
    const excluded = manifest.excludedTables as { reasons: unknown[] }[];
    expect(excluded).toHaveLength(1);
    expect(excluded[0]!.reasons).toEqual([
      { code: "PBI_TABLE_HIDDEN", count: 1 },
    ]);
  }, 120_000);

  it("writes byte-identical workbooks on two runs in one process", async () => {
    const first = await exportPbiTables(fixture);
    const second = await exportPbiTables(fixture);
    expect(Buffer.from(second.outputs[0]!.bytes)).toEqual(
      Buffer.from(first.outputs[0]!.bytes),
    );
    expect(Buffer.from(second.outputs[1]!.bytes)).toEqual(
      Buffer.from(first.outputs[1]!.bytes),
    );
  }, 180_000);

  it("writes byte-identical workbooks from a fresh process", async () => {
    const root = path.resolve(
      path.dirname(fileURLToPath(import.meta.url)),
      "..",
    );
    if (!existsSync(path.join(root, "dist", "index.js"))) {
      // The fresh-process arm needs the built package. CI builds before it
      // tests; a local run without a build keeps the same-process arm above.
      process.stdout.write(
        "fresh-process determinism skipped: packages/pbi/dist is not built\n",
      );
      return;
    }
    const printed = execFileSync(
      process.execPath,
      [path.join(root, "test", "export-once.ts")],
      { encoding: "utf8", cwd: root },
    ).trim();
    const here = await exportPbiTables(fixture);
    expect(printed).toBe(
      here.outputs
        .map((output) =>
          createHash("sha256").update(output.bytes).digest("hex"),
        )
        .join(" "),
    );
  }, 300_000);

  it("reads the same tables through the decode-only entry point", async () => {
    const model = await readPbiTables(fixture);
    expect(model.tables.map((table) => table.name)).toEqual([
      "People",
      "Sales",
    ]);
    expect(model.manifest.schemaVersion).toBe(1);
    expect(model.warnings.length).toBeGreaterThan(0);
  }, 120_000);
});

describe("output and peak limits", () => {
  it("refuses when the workbook and the manifest together pass outputBytes", async () => {
    const baseline = await exportPbiTables(fixture);
    const workbook = baseline.outputs[0]!.bytes.byteLength;
    const manifest = baseline.outputs[1]!.bytes.byteLength;

    // Exactly the two artifacts together: inclusive, so this passes.
    const exact = await exportPbiTables(fixture, {
      outputBytes: workbook + manifest,
    });
    expect(exact.outputs).toHaveLength(2);

    for (const [outputBytes, why] of [
      [workbook + manifest - 1, "one byte over"],
      [workbook, "fits alone but not with its manifest"],
    ] as const) {
      let failure: ConsultChimpsError | undefined;
      try {
        await exportPbiTables(fixture, { outputBytes });
      } catch (error) {
        failure = error as ConsultChimpsError;
      }
      expect(failure, why).toBeDefined();
      expect(failure!.code).toBe("PBI_EXPORT_LIMIT_EXCEEDED");
      expect(failure!.details).toMatchObject({
        stage: "workbook",
        option: "outputBytes",
        limit: outputBytes,
        required: workbook + manifest,
      });
      expect(failure!.cause).toBeUndefined();
    }
  }, 300_000);

  const large = path.join(FETCHED, "n-2020-11.pbix");
  it.skipIf(!existsSync(large))(
    "refuses at the decode stage before it allocates a column",
    async () => {
      // Above what the container, the decompressor and the catalog need, below
      // what the catalog's declared row counts say the decoded columns will
      // cost. The reserve is computed from the catalog, so the refusal happens
      // instead of the allocation rather than after it.
      let failure: ConsultChimpsError | undefined;
      try {
        await readPbiTables(new Uint8Array(readFileSync(large)), {
          peakBytes: 120_000_000,
        });
      } catch (error) {
        failure = error as ConsultChimpsError;
      }
      expect(failure).toBeDefined();
      expect(failure!.code).toBe("PBI_EXPORT_LIMIT_EXCEEDED");
      expect(failure!.details).toMatchObject({
        stage: "decode",
        option: "peakBytes",
        limit: 120_000_000,
      });
      const required = (failure!.details as { required: number }).required;
      expect(required).toBeGreaterThan(120_000_000);
    },
    300_000,
  );

  it("records nothing when a reservation is refused", () => {
    const budget = new PipelineBudget(
      validateExportOptions({ peakBytes: 1000 }, true),
    );
    budget.reserve("first", 900, "decode");
    expect(() => budget.reserve("second", 200, "decode")).toThrow(
      ConsultChimpsError,
    );
    // The refused term never enters the accounting, so a caller that catches
    // and continues cannot be charged for memory it was not allowed to take.
    expect(budget.terms()).toEqual({ first: 900 });
    expect(budget.peakBytes).toBe(900);
  });
});
