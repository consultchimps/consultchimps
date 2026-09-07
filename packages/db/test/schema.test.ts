import { describe, expect, it } from "vitest";

import {
  Database,
  loadSqlDatabase,
  quoteIdentifier,
  sameIdentifier,
  METADATA_TABLE,
  TABLE_REGISTRY_TABLE,
  type TableSchema,
} from "../src/index.js";

const customer: TableSchema = {
  name: "Customer",
  columns: [
    { name: "name", type: "text", nullable: false },
    { name: "active", type: "boolean" },
  ],
  foreignKeys: [],
  recordId: { prefix: "CUST", padding: 4 },
};

const invoice: TableSchema = {
  name: "Invoice",
  columns: [
    { name: "customer", type: "text" },
    { name: "amount", type: "real" },
  ],
  foreignKeys: [{ column: "customer", referencesTable: "Customer" }],
  recordId: { prefix: "INV", padding: 5 },
};

describe("schema round trip through save and load", () => {
  it("restores the schema, id counters, and rows from serialized bytes", async () => {
    const database = await Database.create();
    database.createTable(customer);
    database.createTable(invoice);

    const north = database.insertRecord("Customer", {
      name: "North",
      active: true,
    });
    database.insertRecord("Invoice", {
      customer: north.recordId,
      amount: 1250.5,
    });

    const bytes = database.serialize();
    database.close();

    const reopened = await Database.open(bytes);

    expect(reopened.schemaFormatVersion()).toBe(1);
    expect(reopened.getSchema()).toEqual([customer, invoice]);
    expect(reopened.getTableSchema("Invoice").recordId).toEqual({
      prefix: "INV",
      padding: 5,
    });

    // The counter survives the round trip: the next id continues the sequence.
    const next = reopened.insertRecord("Customer", {
      name: "South",
      active: false,
    });
    expect(next.recordId).toBe("CUST-0002");

    const rows = reopened.readRecords("Customer");
    expect(rows).toEqual([
      { record_id: "CUST-0001", name: "North", active: true },
      { record_id: "CUST-0002", name: "South", active: false },
    ]);
    reopened.close();
  });

  it("enforces a foreign key with a structured error", async () => {
    const database = await Database.create();
    database.createTable(customer);
    database.createTable(invoice);

    expect(() =>
      database.insertRecord("Invoice", {
        customer: "CUST-0404",
        amount: 10,
      }),
    ).toThrow(expect.objectContaining({ code: "DB_FOREIGN_KEY_VIOLATION" }));
    database.close();
  });

  it("reports a missing required column with a structured error", async () => {
    const database = await Database.create();
    database.createTable(customer);
    // "name" is nullable: false and is not supplied.
    expect(() => database.insertRecord("Customer", { active: true })).toThrow(
      expect.objectContaining({ code: "DB_NOT_NULL_VIOLATION" }),
    );
    database.close();
  });

  it("rejects non-numeric text for a numeric column and keeps blanks null", async () => {
    const database = await Database.create();
    database.createTable(customer);
    database.createTable(invoice);
    const north = database.insertRecord("Customer", {
      name: "North",
      active: true,
    });

    expect(() =>
      database.insertRecord("Invoice", {
        customer: north.recordId,
        amount: "not available",
      }),
    ).toThrow(expect.objectContaining({ code: "DB_INVALID_NUMBER" }));

    database.insertRecord("Invoice", {
      customer: north.recordId,
      amount: "",
    });
    expect(database.readRecords("Invoice")).toEqual([
      { record_id: "INV-00001", customer: north.recordId, amount: null },
    ]);
    database.close();
  });

  it("rejects a fractional value for an integer column", async () => {
    const database = await Database.create();
    database.createTable({
      name: "Ticket",
      columns: [{ name: "priority", type: "integer" }],
      foreignKeys: [],
      recordId: { prefix: "TK", padding: 3 },
    });
    expect(() => database.insertRecord("Ticket", { priority: 1.9 })).toThrow(
      expect.objectContaining({ code: "DB_INVALID_NUMBER" }),
    );
    expect(() => database.insertRecord("Ticket", { priority: "2.5" })).toThrow(
      expect.objectContaining({ code: "DB_INVALID_NUMBER" }),
    );
    // A whole number, including whole-valued text, is accepted.
    expect(database.insertRecord("Ticket", { priority: "3" }).recordId).toBe(
      "TK-001",
    );
    database.close();
  });

  it("rejects an insert that gives one column twice under different casing", async () => {
    const database = await Database.create();
    database.createTable(customer);
    expect(() =>
      database.insertRecord("Customer", { name: "a", Name: "b" }),
    ).toThrow(expect.objectContaining({ code: "DB_DUPLICATE_INSERT_COLUMN" }));
    database.close();
  });

  it("rejects a damaged Record ID counter before generating an id", async () => {
    const database = await Database.create();
    database.createTable(customer);
    database.sql.run(
      `UPDATE ${quoteIdentifier(TABLE_REGISTRY_TABLE)} SET next_counter = ? WHERE name = ?;`,
      [0, "Customer"],
    );
    expect(() =>
      database.insertRecord("Customer", { name: "North", active: true }),
    ).toThrow(expect.objectContaining({ code: "DB_CORRUPT_RECORD_COUNTER" }));
    database.close();
  });

  it("matches insert column names case-insensitively", async () => {
    const database = await Database.create();
    database.createTable(customer);
    const inserted = database.insertRecord("Customer", {
      NAME: "North",
      Active: true,
    });
    expect(inserted.recordId).toBe("CUST-0001");
    // Stored and read back under the declared column names.
    expect(database.readRecords("Customer")).toEqual([
      { record_id: "CUST-0001", name: "North", active: true },
    ]);
    database.close();
  });

  it("rejects a boolean number that is not 0 or 1", async () => {
    const database = await Database.create();
    database.createTable(customer);
    expect(() =>
      database.insertRecord("Customer", { name: "X", active: 2 }),
    ).toThrow(expect.objectContaining({ code: "DB_INVALID_BOOLEAN" }));
    // 0 and 1 remain valid.
    expect(
      database.insertRecord("Customer", { name: "Y", active: 1 }).recordId,
    ).toBe("CUST-0001");
    database.close();
  });

  it("accepts a self-referencing foreign key whose table name differs in case", async () => {
    const database = await Database.create();
    database.createTable({
      name: "Node",
      columns: [{ name: "parent", type: "text" }],
      foreignKeys: [{ column: "parent", referencesTable: "node" }],
      recordId: { prefix: "ND", padding: 3 },
    });
    const root = database.insertRecord("Node", { parent: null });
    const child = database.insertRecord("Node", { parent: root.recordId });
    expect(child.recordId).toBe("ND-002");
    database.close();
  });

  it("rejects an integer outside the range representable exactly", async () => {
    const database = await Database.create();
    database.createTable({
      name: "Big",
      columns: [{ name: "n", type: "integer" }],
      foreignKeys: [],
      recordId: { prefix: "B", padding: 2 },
    });
    expect(() =>
      database.insertRecord("Big", { n: "9007199254740993" }),
    ).toThrow(expect.objectContaining({ code: "DB_INVALID_NUMBER" }));
    database.close();
  });

  it("keeps the offending value out of a conversion error but adds context", async () => {
    const database = await Database.create();
    database.createTable(customer);
    let caught:
      | { code: string; message: string; details: Record<string, unknown> }
      | undefined;
    try {
      database.insertRecord("Customer", { name: "X", active: "secret-value" });
    } catch (error) {
      caught = error as typeof caught;
    }
    expect(caught?.code).toBe("DB_INVALID_BOOLEAN");
    expect(caught?.message).not.toContain("secret-value");
    expect(caught?.details).not.toHaveProperty("value");
    expect(caught?.details["table"]).toBe("Customer");
    expect(caught?.details["column"]).toBe("active");
    database.close();
  });

  it("detects a duplicate table name case-insensitively", async () => {
    const database = await Database.create();
    database.createTable(customer);
    expect(() =>
      database.createTable({ ...customer, name: "customer" }),
    ).toThrow(expect.objectContaining({ code: "DB_TABLE_EXISTS" }));
    database.close();
  });

  it("reserves the Record ID column case-insensitively", async () => {
    const database = await Database.create();
    expect(() =>
      database.createTable({
        name: "Thing",
        columns: [{ name: "RECORD_ID", type: "text" }],
        foreignKeys: [],
        recordId: { prefix: "T", padding: 2 },
      }),
    ).toThrow(expect.objectContaining({ code: "DB_RESERVED_COLUMN" }));
    database.close();
  });

  it("refuses the __proto__ identifier that cannot be stored as a data key", async () => {
    const database = await Database.create();
    expect(() =>
      database.createTable({
        name: "Thing",
        columns: [{ name: "__proto__", type: "text" }],
        foreignKeys: [],
        recordId: { prefix: "T", padding: 2 },
      }),
    ).toThrow(expect.objectContaining({ code: "DB_RESERVED_IDENTIFIER" }));
    database.close();
  });

  it("handles a column named like an Object.prototype member safely", async () => {
    const database = await Database.create();
    database.createTable({
      name: "Thing",
      columns: [{ name: "constructor", type: "text" }],
      foreignKeys: [],
      recordId: { prefix: "T", padding: 2 },
    });
    database.insertRecord("Thing", { constructor: "value" });
    // A record whose row object lacks its own "constructor" value stores null,
    // not the inherited Object constructor.
    database.insertRecord("Thing", {});
    expect(
      database.readRecords("Thing").map((row) => row["constructor"]),
    ).toEqual(["value", null]);
    database.close();
  });

  it("rejects the SQLite-reserved table name prefix", async () => {
    const database = await Database.create();
    expect(() =>
      database.createTable({
        name: "sqlite_stat1",
        columns: [{ name: "a", type: "text" }],
        foreignKeys: [],
        recordId: { prefix: "S", padding: 2 },
      }),
    ).toThrow(/sqlite_/i);
    database.close();
  });

  it("rejects a duplicate table and a reserved column name", async () => {
    const database = await Database.create();
    database.createTable(customer);
    expect(() => database.createTable(customer)).toThrow(/already exists/i);
    expect(() =>
      database.createTable({
        name: "Bad",
        columns: [{ name: "record_id", type: "text" }],
        foreignKeys: [],
        recordId: { prefix: "B", padding: 2 },
      }),
    ).toThrow(/reserved/i);
    database.close();
  });

  it("rejects a foreign key to a missing table", async () => {
    const database = await Database.create();
    expect(() =>
      database.createTable({
        name: "Orphan",
        columns: [{ name: "parent", type: "text" }],
        foreignKeys: [{ column: "parent", referencesTable: "Nowhere" }],
        recordId: { prefix: "O", padding: 3 },
      }),
    ).toThrow(/does not exist/i);
    database.close();
  });

  it("coerces boolean text explicitly rather than by truthiness", async () => {
    const database = await Database.create();
    database.createTable(customer);

    database.insertRecord("Customer", { name: "A", active: "false" });
    database.insertRecord("Customer", { name: "B", active: "0" });
    database.insertRecord("Customer", { name: "C", active: "yes" });

    expect(
      database.readRecords("Customer").map((row) => row["active"]),
    ).toEqual([false, false, true]);

    expect(() =>
      database.insertRecord("Customer", { name: "D", active: "maybe" }),
    ).toThrow(/boolean/i);
    database.close();
  });

  it("rejects a database written with a newer schema format version", async () => {
    const database = await Database.create();
    database.createTable(customer);
    database.sql.run(
      `UPDATE ${quoteIdentifier(METADATA_TABLE)} SET value = ? WHERE key = ?;`,
      ["99", "schema_format_version"],
    );
    const bytes = database.serialize();
    database.close();

    await expect(Database.open(bytes)).rejects.toThrow(/newer/i);
  });

  it("rejects a database whose schema version row is missing or malformed", async () => {
    const database = await Database.create();
    database.createTable(customer);
    database.sql.run(
      `DELETE FROM ${quoteIdentifier(METADATA_TABLE)} WHERE key = ?;`,
      ["schema_format_version"],
    );
    const bytes = database.serialize();
    database.close();

    await expect(Database.open(bytes)).rejects.toThrow(
      expect.objectContaining({ code: "DB_UNSUPPORTED_SCHEMA_VERSION" }),
    );
  });

  it("rejects a corrupt stored value when reading records", async () => {
    const database = await Database.create();
    database.createTable(customer);
    database.insertRecord("Customer", { name: "North", active: true });
    // Corrupt the stored boolean through the raw handle, as an external edit
    // could.
    database.sql.run(
      `UPDATE ${quoteIdentifier("Customer")} SET ${quoteIdentifier("active")} = 2;`,
    );
    expect(() => database.readRecords("Customer")).toThrow(
      expect.objectContaining({ code: "DB_CORRUPT_STORED_VALUE" }),
    );
    database.close();
  });

  it("rejects opening a database with a malformed stored definition", async () => {
    const database = await Database.create();
    database.createTable(customer);
    database.sql.run(
      `UPDATE ${quoteIdentifier(TABLE_REGISTRY_TABLE)} SET definition = ? WHERE name = ?;`,
      ["{not json", "Customer"],
    );
    const bytes = database.serialize();
    database.close();
    await expect(Database.open(bytes)).rejects.toThrow(
      expect.objectContaining({ code: "DB_CORRUPT_WORKSPACE" }),
    );
  });

  it("returns the declared table name even when looked up in another case", async () => {
    const database = await Database.create();
    database.createTable(customer);
    expect(database.getTableSchema("customer").name).toBe("Customer");
    database.close();
  });

  it("folds identifiers like SQLite NOCASE: ASCII case only", () => {
    expect(sameIdentifier("Customer", "customer")).toBe(true);
    expect(sameIdentifier("ABC", "abc")).toBe(true);
    // Non-ASCII case is not folded, matching SQLite's NOCASE.
    expect(sameIdentifier("Ä", "ä")).toBe(false);
  });

  it("refuses to open bytes that are not a ConsultChimps database", async () => {
    // A raw sql.js database with no metadata tables.
    const raw = await loadSqlDatabase();
    raw.run("CREATE TABLE loose (a INTEGER);");
    const bytes = raw.serialize();
    raw.close();

    await expect(Database.open(bytes)).rejects.toThrow(/ConsultChimps/i);
  });
});
