import { describe, expect, it } from "vitest";

import type { CellValue } from "@consultchimps/tabular";

import {
  Database,
  RECORD_ID_COLUMN,
  updateRecords,
  type TableSchema,
} from "../src/index.js";

// updateRecords is the batch behind one spreadsheet gesture: a paste or a fill
// is many cell writes that must reach the file as one step. These tests pin the
// two halves of that promise. Every accepted write commits together, and a
// value the schema refuses is reported against its own request rather than
// discarding the writes beside it.

const region: TableSchema = {
  name: "Region",
  columns: [{ name: "name", type: "text", nullable: false }],
  foreignKeys: [],
  recordId: { prefix: "REG", padding: 4 },
};

const customer: TableSchema = {
  name: "Customer",
  columns: [
    { name: "name", type: "text", nullable: false },
    { name: "region", type: "text" },
    { name: "headcount", type: "integer" },
    { name: "active", type: "boolean" },
  ],
  foreignKeys: [{ column: "region", referencesTable: "Region" }],
  recordId: { prefix: "CUST", padding: 4 },
};

/** Two regions and two customers, enough for a small rectangular gesture. */
async function seeded(): Promise<Database> {
  const database = await Database.create();
  database.createTable(region);
  database.createTable(customer);
  database.insertRecord("Region", { name: "North" });
  database.insertRecord("Region", { name: "South" });
  database.insertRecord("Customer", {
    name: "Acme",
    region: "REG-0001",
    headcount: 12,
    active: true,
  });
  database.insertRecord("Customer", {
    name: "Globex",
    region: "REG-0002",
    headcount: 7,
    active: false,
  });
  return database;
}

function stored(
  database: Database,
  table: string,
  recordId: string,
  column: string,
): CellValue | undefined {
  const row = database
    .readRecords(table)
    .find((record) => record[RECORD_ID_COLUMN] === recordId);
  return row?.[column];
}

describe("updateRecords", () => {
  it("applies every accepted write and reports what the database stored", async () => {
    const database = await seeded();

    const outcomes = updateRecords(database, [
      { table: "Customer", recordId: "CUST-0001", values: { headcount: "20" } },
      { table: "Customer", recordId: "CUST-0002", values: { headcount: "21" } },
    ]);

    expect(outcomes).toEqual([
      { accepted: true, recordId: "CUST-0001", values: { headcount: 20 } },
      { accepted: true, recordId: "CUST-0002", values: { headcount: 21 } },
    ]);
    expect(stored(database, "Customer", "CUST-0001", "headcount")).toBe(20);
    expect(stored(database, "Customer", "CUST-0002", "headcount")).toBe(21);
    database.close();
  });

  it("reads a converted value back rather than the value it was sent", async () => {
    const database = await seeded();

    const outcomes = updateRecords(database, [
      { table: "Customer", recordId: "CUST-0002", values: { active: "yes" } },
    ]);

    expect(outcomes).toEqual([
      { accepted: true, recordId: "CUST-0002", values: { active: true } },
    ]);
    database.close();
  });

  it("refuses one impossible value and keeps the writes beside it", async () => {
    const database = await seeded();

    const outcomes = updateRecords(database, [
      { table: "Customer", recordId: "CUST-0001", values: { headcount: "30" } },
      {
        table: "Customer",
        recordId: "CUST-0002",
        values: { headcount: "1.5" },
      },
    ]);

    expect(outcomes[0]?.accepted).toBe(true);
    expect(outcomes[1]).toMatchObject({
      accepted: false,
      code: "DB_INVALID_NUMBER",
      recordId: "CUST-0002",
    });
    expect(stored(database, "Customer", "CUST-0001", "headcount")).toBe(30);
    // The refused write changed nothing, and the accepted one beside it stands.
    expect(stored(database, "Customer", "CUST-0002", "headcount")).toBe(7);
    database.close();
  });

  it("reports a foreign key that names no record, without losing the rest", async () => {
    const database = await seeded();

    const outcomes = updateRecords(database, [
      {
        table: "Customer",
        recordId: "CUST-0001",
        values: { region: "REG-9999" },
      },
      { table: "Customer", recordId: "CUST-0002", values: { name: "Initech" } },
    ]);

    expect(outcomes[0]?.accepted).toBe(false);
    expect(outcomes[1]?.accepted).toBe(true);
    expect(stored(database, "Customer", "CUST-0001", "region")).toBe(
      "REG-0001",
    );
    expect(stored(database, "Customer", "CUST-0002", "name")).toBe("Initech");
    database.close();
  });

  it("reports an unknown record, column, table, and Record ID write per request", async () => {
    const database = await seeded();

    const outcomes = updateRecords(database, [
      { table: "Customer", recordId: "CUST-9999", values: { name: "Ghost" } },
      { table: "Customer", recordId: "CUST-0001", values: { nope: "1" } },
      { table: "Nowhere", recordId: "CUST-0001", values: { name: "x" } },
      {
        table: "Customer",
        recordId: "CUST-0001",
        values: { [RECORD_ID_COLUMN]: "CUST-0009" },
      },
    ]);

    expect(outcomes.map((outcome) => outcome.accepted)).toEqual([
      false,
      false,
      false,
      false,
    ]);
    expect(
      outcomes.map((outcome) => (outcome.accepted ? null : outcome.code)),
    ).toEqual([
      "DB_RECORD_NOT_FOUND",
      "DB_UNKNOWN_COLUMN",
      "DB_TABLE_NOT_FOUND",
      "DB_RECORD_ID_IMMUTABLE",
    ]);
    database.close();
  });

  it("writes several columns of one record when they are asked for together", async () => {
    const database = await seeded();

    const outcomes = updateRecords(database, [
      {
        table: "Customer",
        recordId: "CUST-0001",
        values: { headcount: "44", name: "Acme Group" },
      },
    ]);

    expect(outcomes[0]?.accepted).toBe(true);
    expect(stored(database, "Customer", "CUST-0001", "name")).toBe(
      "Acme Group",
    );
    expect(stored(database, "Customer", "CUST-0001", "headcount")).toBe(44);
    database.close();
  });

  it("writes across two tables in the one batch", async () => {
    const database = await seeded();

    const outcomes = updateRecords(database, [
      { table: "Region", recordId: "REG-0001", values: { name: "Upper" } },
      { table: "Customer", recordId: "CUST-0001", values: { name: "Initech" } },
    ]);

    expect(outcomes.every((outcome) => outcome.accepted)).toBe(true);
    expect(stored(database, "Region", "REG-0001", "name")).toBe("Upper");
    expect(stored(database, "Customer", "CUST-0001", "name")).toBe("Initech");
    database.close();
  });

  it("rolls every write back when a fault that is not a refusal escapes", async () => {
    const database = await seeded();
    // A refusal is a ConsultChimpsError and belongs to its own request. Anything
    // else is a fault nobody planned for, and the batch has to come back off
    // rather than commit the half of it that ran first. Reading the values
    // throws here, which is the plainest way to produce one.
    const faulty = new Proxy({} as Record<string, CellValue>, {
      ownKeys: () => ["name"],
      getOwnPropertyDescriptor: () => ({
        configurable: true,
        enumerable: true,
        value: "x",
        writable: true,
      }),
      get: () => {
        throw new Error("Something unplanned happened while reading a value.");
      },
    });

    expect(() =>
      updateRecords(database, [
        { table: "Customer", recordId: "CUST-0001", values: { headcount: 5 } },
        { table: "Customer", recordId: "CUST-0002", values: faulty },
      ]),
    ).toThrow("Something unplanned happened");

    expect(stored(database, "Customer", "CUST-0001", "headcount")).toBe(12);
    database.close();
  });

  it("rolls every write back when the workspace turns out to be damaged", async () => {
    const database = await seeded();
    // A refusal is about a request. A corruption error is a library error too,
    // but it is about the file, found by whichever request happened to reach
    // it, and the writes beside it must not commit into a file known to be
    // broken. The registry is damaged here the way an advanced caller with the
    // raw handle could damage it.
    database.sql.run(
      `UPDATE "_consultchimps_tables" SET definition = 'not json' WHERE name = 'Customer';`,
    );

    expect(() =>
      updateRecords(database, [
        {
          table: "Region",
          recordId: "REG-0001",
          values: { name: "Far North" },
        },
        { table: "Customer", recordId: "CUST-0001", values: { headcount: 5 } },
      ]),
    ).toThrow(expect.objectContaining({ code: "DB_CORRUPT_WORKSPACE" }));

    expect(stored(database, "Region", "REG-0001", "name")).toBe("North");
    database.close();
  });

  it("writes nothing and reports nothing for an empty batch", async () => {
    const database = await seeded();

    expect(updateRecords(database, [])).toEqual([]);
    expect(stored(database, "Customer", "CUST-0001", "headcount")).toBe(12);
    database.close();
  });

  it("carries the refusal's own sentence, so a caller can show it", async () => {
    const database = await seeded();

    const outcomes = updateRecords(database, [
      {
        table: "Customer",
        recordId: "CUST-0001",
        values: { headcount: "many" },
      },
    ]);

    const outcome = outcomes[0];
    expect(outcome?.accepted).toBe(false);
    expect(outcome?.accepted === false && outcome.message).toContain(
      "whole number",
    );
    expect(outcome?.accepted === false && outcome.message).toContain(
      "headcount",
    );
    database.close();
  });

  it("leaves an accepted batch durable in the serialized bytes", async () => {
    const database = await seeded();

    updateRecords(database, [
      { table: "Customer", recordId: "CUST-0001", values: { headcount: "99" } },
    ]);
    const bytes = database.serialize();
    database.close();

    const reopened = await Database.open(bytes);
    expect(stored(reopened, "Customer", "CUST-0001", "headcount")).toBe(99);
    reopened.close();
  });
});
