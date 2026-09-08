import { describe, expect, it } from "vitest";

import {
  Database,
  formatRecordId,
  quoteIdentifier,
  type TableSchema,
} from "../src/index.js";

function customerSchema(): TableSchema {
  return {
    name: "Customer",
    columns: [{ name: "name", type: "text" }],
    foreignKeys: [],
    recordId: { prefix: "CUST", padding: 4 },
  };
}

describe("formatRecordId", () => {
  it("applies the prefix, separator, and zero padding", () => {
    expect(formatRecordId({ prefix: "CUST", padding: 4 }, 1)).toBe("CUST-0001");
    expect(formatRecordId({ prefix: "INV", padding: 4 }, 42)).toBe("INV-0042");
    expect(formatRecordId({ prefix: "X", padding: 2, separator: "_" }, 7)).toBe(
      "X_07",
    );
    expect(formatRecordId({ prefix: "N", padding: 0 }, 123)).toBe("N-123");
  });
});

describe("generated Record IDs", () => {
  it("assigns a sequential id to each inserted record", async () => {
    const database = await Database.create();
    database.createTable(customerSchema());

    const first = database.insertRecord("Customer", { name: "North" });
    const second = database.insertRecord("Customer", { name: "South" });
    const third = database.insertRecord("Customer", { name: "East" });

    expect(first.recordId).toBe("CUST-0001");
    expect(second.recordId).toBe("CUST-0002");
    expect(third.recordId).toBe("CUST-0003");
    database.close();
  });

  it("does not reuse an id after a delete, so gaps are fine", async () => {
    const database = await Database.create();
    database.createTable(customerSchema());

    database.insertRecord("Customer", { name: "North" });
    database.insertRecord("Customer", { name: "South" });
    database.insertRecord("Customer", { name: "East" });

    database.sql.run(
      `DELETE FROM ${quoteIdentifier("Customer")} WHERE ${quoteIdentifier("record_id")} = ?;`,
      ["CUST-0002"],
    );

    const next = database.insertRecord("Customer", { name: "West" });
    expect(next.recordId).toBe("CUST-0004");

    const ids = database.readRecords("Customer").map((row) => row["record_id"]);
    expect(ids).toEqual(["CUST-0001", "CUST-0003", "CUST-0004"]);
    database.close();
  });

  it("rejects supplying a Record ID at insert time", async () => {
    const database = await Database.create();
    database.createTable(customerSchema());
    expect(() =>
      database.insertRecord("Customer", { record_id: "CUST-9999", name: "X" }),
    ).toThrow(/generated/i);
    database.close();
  });

  it("refuses to change a Record ID once assigned", async () => {
    const database = await Database.create();
    database.createTable(customerSchema());
    database.insertRecord("Customer", { name: "North" });

    expect(() =>
      database.sql.run(
        `UPDATE ${quoteIdentifier("Customer")} SET ${quoteIdentifier("record_id")} = ? WHERE ${quoteIdentifier("record_id")} = ?;`,
        ["CUST-0002", "CUST-0001"],
      ),
    ).toThrow(/immutable/i);
    database.close();
  });
});
