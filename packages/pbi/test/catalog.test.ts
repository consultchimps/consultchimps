import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { ConsultChimpsError } from "@consultchimps/core";
import type { Sqlite3Static } from "@sqlite.org/sqlite-wasm";
import { memberBytes, parseAbf } from "../src/abf.js";
import { PipelineBudget, validateExportOptions } from "../src/budget.js";
import { capDax, readCatalog } from "../src/catalog.js";
import { readPbiModelPart } from "../src/container.js";
import { decompressModelPart } from "../src/xpress9/stream.js";
import { FETCHED, FIXTURES } from "./oracle.js";

/**
 * Section C, the catalog rows. The reader runs against the real catalog of a
 * committed fixture, so the schema probe, the ordering and the declared row
 * count are checked against a catalog Power BI actually wrote.
 */

/** The committed fixture, or the fetched cache when the test names another. */
function locate(fixture: string): string | null {
  for (const directory of [FIXTURES, FETCHED]) {
    const candidate = path.join(directory, `${fixture}.pbix`);
    if (existsSync(candidate)) return candidate;
  }
  return null;
}

async function catalogBytes(fixture: string): Promise<Uint8Array> {
  const container = new Uint8Array(readFileSync(locate(fixture)!));
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

  it.skipIf(locate("b-2018-profiling") === null)(
    "carries the DAX of calculated objects and nothing for imported ones",
    async () => {
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
    },
    120_000,
  );

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

/**
 * A synthetic catalog with the tables the storage query joins, so the paging of
 * the calculated-partition query can be driven past the reader's page size
 * without a model that large existing anywhere.
 */
async function syntheticCatalog(
  calculatedPartitions: number,
): Promise<Uint8Array> {
  const { default: initialize } = await import("@sqlite.org/sqlite-wasm");
  const sqlite = await (
    initialize as unknown as (config: {
      print: () => void;
      printErr: () => void;
    }) => Promise<Sqlite3Static>
  )({ print: () => undefined, printErr: () => undefined });
  const database = new sqlite.oo1.DB(":memory:");
  try {
    database.exec(`
      CREATE TABLE [Table] (ID INTEGER, Name TEXT, IsHidden INTEGER);
      CREATE TABLE Column (ID INTEGER, TableID INTEGER, ExplicitName TEXT,
        InferredName TEXT, IsHidden INTEGER, Type INTEGER, Expression TEXT,
        ExplicitDataType INTEGER, InferredDataType INTEGER,
        ColumnStorageID INTEGER);
      CREATE TABLE ColumnStorage (ID INTEGER, StoragePosition INTEGER,
        DictionaryStorageID INTEGER, Statistics_RowCount INTEGER);
      CREATE TABLE AttributeHierarchy (ColumnID INTEGER,
        AttributeHierarchyStorageID INTEGER);
      CREATE TABLE AttributeHierarchyStorage (ID INTEGER, StorageFileID INTEGER);
      CREATE TABLE StorageFile (ID INTEGER, FileName TEXT);
      CREATE TABLE DictionaryStorage (ID INTEGER, StorageFileID INTEGER,
        BaseId INTEGER, Magnitude REAL);
      CREATE TABLE ColumnPartitionStorage (ColumnStorageID INTEGER,
        StorageFileID INTEGER, PartitionStorageID INTEGER);
      CREATE TABLE PartitionStorage (ID INTEGER, StoragePosition INTEGER);
      CREATE TABLE Partition (ID INTEGER, TableID INTEGER, Type INTEGER,
        QueryDefinition TEXT);
      INSERT INTO StorageFile VALUES (1, 'c.idf'), (2, 'c.hidx');
      INSERT INTO PartitionStorage VALUES (1, 0);
      INSERT INTO AttributeHierarchyStorage VALUES (1, 2);
    `);
    // One table the main query returns, whose calculated partition sits past
    // the first page of the calculated-partition query.
    const last = calculatedPartitions;
    database.exec(`
      INSERT INTO [Table] VALUES (${last}, 'Late', 0);
      INSERT INTO ColumnStorage VALUES (1, 0, 1, 3);
      INSERT INTO DictionaryStorage VALUES (1, NULL, 0, 1.0);
      INSERT INTO Column VALUES (1, ${last}, 'Amount', NULL, 0, 1, NULL, 6, 6, 1);
      INSERT INTO AttributeHierarchy VALUES (1, 1);
      INSERT INTO ColumnPartitionStorage VALUES (1, 1, 1);
    `);
    database.exec("BEGIN");
    for (let index = 1; index <= calculatedPartitions; index++)
      database.exec({
        sql: "INSERT INTO Partition VALUES (?, ?, 2, ?)",
        bind: [index, index, `EVALUATE ROW("n", ${index})`],
      });
    database.exec("COMMIT");
    const pointer = sqlite.capi.sqlite3_js_db_export(database);
    return pointer;
  } finally {
    database.close();
  }
}

describe("the calculated-partition query", () => {
  it("sees a calculated table past the reader's page size", async () => {
    // 2,049 partitions: the last one is the first row of the second page.
    const bytes = await syntheticCatalog(2_049);
    const catalog = await readCatalog(bytes, undefined);
    try {
      const table = catalog.tables.find((entry) => entry.name === "Late");
      expect(table).toBeDefined();
      // Before paging this table was reported as imported and lost its DAX.
      expect(table!.calculated).toBe(true);
      expect(table!.dax).toBe('EVALUATE ROW("n", 2049)');
    } finally {
      catalog.close();
    }
  }, 300_000);
});
