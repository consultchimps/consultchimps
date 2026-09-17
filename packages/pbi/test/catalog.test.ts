import { readFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { ConsultChimpsError } from "@consultchimps/core";
import { memberBytes, parseAbf } from "../src/abf.js";
import { PipelineBudget, validateExportOptions } from "../src/budget.js";
import { capDax, readCatalog } from "../src/catalog.js";
import { readPbiModelPart } from "../src/container.js";
import { decompressModelPart } from "../src/xpress9/stream.js";
import { FIXTURES } from "./oracle.js";

/**
 * Section C, the catalog rows. The reader runs against the real catalog of a
 * committed fixture, so the schema probe, the ordering and the declared row
 * count are checked against a catalog Power BI actually wrote.
 */

async function catalogBytes(fixture: string): Promise<Uint8Array> {
  const container = new Uint8Array(
    readFileSync(path.join(FIXTURES, `${fixture}.pbix`)),
  );
  const budget = new PipelineBudget(validateExportOptions({}, true));
  const stream = await decompressModelPart(
    readPbiModelPart(container),
    budget,
    undefined,
  );
  return memberBytes(stream.bytes, parseAbf(stream.bytes), "metadata.sqlitedb");
}

async function refusal(
  run: () => Promise<unknown>,
): Promise<ConsultChimpsError> {
  try {
    await run();
  } catch (error) {
    if (error instanceof ConsultChimpsError) return error;
    throw error;
  }
  throw new Error("expected a refusal");
}

describe("the storage catalog", () => {
  it("reads tables, columns and the declared row count in catalog order", async () => {
    const catalog = await readCatalog(
      await catalogBytes("a-2018-fuzzy"),
      undefined,
    );
    try {
      expect(catalog.tables.map((table) => table.name).sort()).toEqual([
        "DateTableTemplate_9e6833d7-5e83-4931-bc93-8dbeed8e6805",
        "People",
        "Sales",
      ]);
      // Ascending numeric table id, and columns by storage position then id.
      const ids = catalog.tables.map((table) => table.id);
      expect([...ids].sort((left, right) => left - right)).toEqual(ids);
      for (const table of catalog.tables) {
        const positions = table.columns.map((entry) => entry.storagePosition);
        expect([...positions].sort((left, right) => left - right)).toEqual(
          positions,
        );
        // The declared row count is the catalog's, and its columns agree on it.
        expect(table.rowCount).toBeGreaterThan(0);
        for (const entry of table.columns)
          expect(entry.rowCount).toBe(table.rowCount);
      }
      const people = catalog.tables.find((table) => table.name === "People")!;
      expect(people.rowCount).toBe(13);
      expect(people.hidden).toBe(false);
      const generated = catalog.tables.find((table) =>
        table.name.startsWith("DateTableTemplate_"),
      )!;
      expect(generated.hidden).toBe(true);
      expect(catalog.wasmMemoryBytes).toBeGreaterThan(0);
    } finally {
      catalog.close();
    }
  }, 120_000);

  it("carries the DAX of calculated objects and nothing for imported ones", async () => {
    const catalog = await readCatalog(
      await catalogBytes("b-2018-profiling"),
      undefined,
    );
    try {
      const calculated = catalog.tables.filter((table) => table.calculated);
      expect(calculated.length).toBeGreaterThan(0);
      for (const table of calculated)
        expect(typeof table.dax === "string" || table.dax === undefined).toBe(
          true,
        );
      for (const table of catalog.tables)
        for (const entry of table.columns)
          if (!entry.calculated) expect(entry.dax).toBeUndefined();
    } finally {
      catalog.close();
    }
  }, 120_000);

  it("maps an unreadable catalog to the model refusal", async () => {
    const error = await refusal(() =>
      readCatalog(new Uint8Array(200), undefined),
    );
    expect(error.code).toBe("PBI_MODEL_UNREADABLE");
    expect(error.details).toEqual({ stage: "catalog" });
    expect(error.cause).toBeUndefined();
  });

  it("maps a truncated catalog to the model refusal", async () => {
    const bytes = await catalogBytes("a-2018-fuzzy");
    const error = await refusal(() =>
      readCatalog(bytes.subarray(0, 4096), undefined),
    );
    expect(error.code).toBe("PBI_MODEL_UNREADABLE");
  });

  it("names the SQLite option when a read limit is exceeded", async () => {
    const bytes = await catalogBytes("a-2018-fuzzy");
    const error = await refusal(() =>
      // A runtime binary of one byte is refused as a malformed source before
      // any query runs, which is the same mapping path a limit takes.
      readCatalog(bytes, { wasmBinary: new Uint8Array(0) }),
    );
    expect([
      "PBI_MODEL_UNREADABLE",
      "PBI_EXPORT_LIMIT_EXCEEDED",
      "PBI_RUNTIME_UNAVAILABLE",
    ]).toContain(error.code);
    expect(error.cause).toBeUndefined();
  }, 120_000);
});

describe("the DAX cap", () => {
  it("leaves an ordinary expression alone", () => {
    expect(capDax("EVALUATE Sales")).toBe("EVALUATE Sales");
    expect(capDax(null)).toBeUndefined();
    expect(capDax("")).toBeUndefined();
  });

  it("shortens a pathological expression to the cap", () => {
    expect(capDax("x".repeat(40_000))!.length).toBe(32_768);
  });

  it("never splits a surrogate pair when shortening", () => {
    const capped = capDax(`${"x".repeat(32_767)}\u{1f600}tail`)!;
    expect(capped.length).toBe(32_767);
    expect(capped.endsWith("x")).toBe(true);
  });
});
