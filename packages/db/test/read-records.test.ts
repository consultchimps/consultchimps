import { isConsultChimpsError } from "@consultchimps/core";
import { describe, expect, it } from "vitest";

import { Database, RECORD_ID_COLUMN } from "../src/index.js";

async function seeded(): Promise<Database> {
  const database = await Database.create();
  database.createTable({
    name: "Customer",
    columns: [
      { name: "name", type: "text" },
      { name: "size", type: "integer" },
    ],
    foreignKeys: [],
    recordId: { prefix: "CUST", padding: 4 },
  });
  for (const [name, size] of [
    ["North", 1],
    ["South", 2],
    ["East", 3],
  ] as const) {
    database.insertRecord("Customer", { name, size });
  }
  return database;
}

function codeOf(run: () => unknown): string {
  try {
    run();
  } catch (error) {
    return isConsultChimpsError(error) ? error.code : "NOT_A_STABLE_ERROR";
  }
  return "NO_ERROR_THROWN";
}

describe("readRecords options", () => {
  it("reads every column and every record by default", async () => {
    const database = await seeded();
    const rows = database.readRecords("Customer");
    expect(rows).toHaveLength(3);
    expect(Object.keys(rows[0] ?? {})).toEqual([
      RECORD_ID_COLUMN,
      "name",
      "size",
    ]);
  });

  it("projects to the requested columns, Record ID first", async () => {
    const database = await seeded();
    const rows = database.readRecords("Customer", { columns: ["size"] });
    expect(rows.map((row) => Object.keys(row))).toEqual([
      [RECORD_ID_COLUMN, "size"],
      [RECORD_ID_COLUMN, "size"],
      [RECORD_ID_COLUMN, "size"],
    ]);
    expect(rows[2]).toEqual({ [RECORD_ID_COLUMN]: "CUST-0003", size: 3 });
  });

  it("matches requested columns like every other identifier", async () => {
    const database = await seeded();
    const [row] = database.readRecords("Customer", { columns: ["NAME"] });
    expect(row).toEqual({ [RECORD_ID_COLUMN]: "CUST-0001", name: "North" });
    // Naming the Record ID changes nothing: it is always read.
    const [again] = database.readRecords("Customer", {
      columns: [RECORD_ID_COLUMN, "name"],
    });
    expect(again).toEqual(row);
  });

  it("limits records in storage order", async () => {
    const database = await seeded();
    expect(
      database
        .readRecords("Customer", { limit: 2 })
        .map((row) => row[RECORD_ID_COLUMN]),
    ).toEqual(["CUST-0001", "CUST-0002"]);
    expect(database.readRecords("Customer", { limit: 0 })).toEqual([]);
    expect(database.readRecords("Customer", { limit: 10 })).toHaveLength(3);
  });

  it("refuses a column the table does not have", async () => {
    const database = await seeded();
    expect(
      codeOf(() => database.readRecords("Customer", { columns: ["region"] })),
    ).toBe("DB_UNKNOWN_COLUMN");
  });

  it("refuses a column named twice", async () => {
    const database = await seeded();
    expect(
      codeOf(() =>
        database.readRecords("Customer", { columns: ["name", "Name"] }),
      ),
    ).toBe("DB_DUPLICATE_READ_COLUMN");
  });

  it("refuses a limit that is not a whole number of zero or more", async () => {
    const database = await seeded();
    for (const limit of [1.5, -1, Number.NaN]) {
      expect(codeOf(() => database.readRecords("Customer", { limit }))).toBe(
        "DB_INVALID_LIMIT",
      );
    }
  });
});
