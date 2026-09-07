import type { Table } from "@consultchimps/tabular";
import { describe, expect, it } from "vitest";

import {
  addRecordsFromTable,
  databaseTableToTable,
  Database,
  type TableSchema,
} from "../src/index.js";

function customerSchema(name: string): TableSchema {
  return {
    name,
    columns: [
      { name: "name", type: "text" },
      { name: "active", type: "boolean" },
      { name: "score", type: "real" },
    ],
    foreignKeys: [],
    recordId: { prefix: "CUST", padding: 4 },
  };
}

describe("databaseTableToTable", () => {
  it("reads a database table out as a Table with coerced values", async () => {
    const database = await Database.create();
    database.createTable(customerSchema("Customer"));
    database.insertRecord("Customer", {
      name: "North",
      active: true,
      score: 12.5,
    });
    database.insertRecord("Customer", {
      name: "South",
      active: false,
      score: 0,
    });

    const table = databaseTableToTable(database, "Customer");
    expect(table.columns).toEqual(["record_id", "name", "active", "score"]);
    expect(table.rows).toEqual([
      { record_id: "CUST-0001", name: "North", active: true, score: 12.5 },
      { record_id: "CUST-0002", name: "South", active: false, score: 0 },
    ]);
    database.close();
  });
});

describe("addRecordsFromTable", () => {
  it("inserts Table rows, generating fresh Record IDs and ignoring an input id column", async () => {
    const database = await Database.create();
    database.createTable(customerSchema("Customer"));

    const input: Table = {
      // A record_id column in the input is ignored; ids are always generated.
      columns: ["record_id", "name", "active", "score"],
      rows: [
        { record_id: "IGNORED-1", name: "East", active: true, score: 3 },
        { record_id: "IGNORED-2", name: "West", active: false, score: 4 },
      ],
    };

    const inserted = addRecordsFromTable(database, "Customer", input);
    expect(inserted.map((record) => record.recordId)).toEqual([
      "CUST-0001",
      "CUST-0002",
    ]);

    const roundTripped = databaseTableToTable(database, "Customer");
    expect(roundTripped.rows).toEqual([
      { record_id: "CUST-0001", name: "East", active: true, score: 3 },
      { record_id: "CUST-0002", name: "West", active: false, score: 4 },
    ]);
    database.close();
  });

  it("round-trips a database table out to a Table and back into a fresh table", async () => {
    const database = await Database.create();
    database.createTable(customerSchema("Source"));
    database.createTable(customerSchema("Target"));

    database.insertRecord("Source", { name: "North", active: true, score: 1 });
    database.insertRecord("Source", { name: "South", active: false, score: 2 });

    const out = databaseTableToTable(database, "Source");
    addRecordsFromTable(database, "Target", out);

    const target = databaseTableToTable(database, "Target");
    expect(target.rows.map((row) => row["name"])).toEqual(["North", "South"]);
    // Record IDs are regenerated on the target, not carried across.
    expect(target.rows.map((row) => row["record_id"])).toEqual([
      "CUST-0001",
      "CUST-0002",
    ]);
    database.close();
  });

  it("rejects a Table column the target table does not declare", async () => {
    const database = await Database.create();
    database.createTable(customerSchema("Customer"));
    const input: Table = {
      columns: ["name", "unknown"],
      rows: [{ name: "North", unknown: "x" }],
    };
    expect(() => addRecordsFromTable(database, "Customer", input)).toThrow(
      /no column/i,
    );
    database.close();
  });
});
