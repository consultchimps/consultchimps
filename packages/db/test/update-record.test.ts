import { ConsultChimpsError } from "@consultchimps/core";
import { describe, expect, it } from "vitest";

import { Database, quoteIdentifier, type TableSchema } from "../src/index.js";

// updateRecord is the one write path the record grid uses, so these tests pin
// what it accepts, what it refuses, and what it reports back. The type rules
// themselves live in coercion.test.ts: what matters here is that an update
// routes through the same conversion an insert does, so neither path can drift
// into accepting a value the other rejects.

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

/** A workspace with two regions and one customer pointing at the first. */
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
  return database;
}

function codeOf(run: () => unknown): string {
  try {
    run();
  } catch (error) {
    return error instanceof ConsultChimpsError ? error.code : "NOT_A_DB_ERROR";
  }
  return "NO_ERROR_THROWN";
}

describe("updateRecord", () => {
  it("writes the given columns and leaves the rest alone", async () => {
    const database = await seeded();

    const updated = database.updateRecord("Customer", "CUST-0001", {
      name: "Acme Holdings",
      headcount: 15,
    });

    expect(updated.recordId).toBe("CUST-0001");
    expect(updated.values).toEqual({ name: "Acme Holdings", headcount: 15 });
    expect(database.readRecords("Customer")).toEqual([
      {
        record_id: "CUST-0001",
        name: "Acme Holdings",
        region: "REG-0001",
        headcount: 15,
        active: true,
      },
    ]);
    database.close();
  });

  it("reports the stored value, not the value it was given", async () => {
    const database = await seeded();

    // The boolean column accepts the text spellings, so what comes back tells
    // the caller which value SQLite actually holds.
    const updated = database.updateRecord("Customer", "CUST-0001", {
      active: "no",
      headcount: "42",
    });

    expect(updated.values).toEqual({ active: false, headcount: 42 });
    database.close();
  });

  it("clears a nullable column when given null", async () => {
    const database = await seeded();

    const updated = database.updateRecord("Customer", "CUST-0001", {
      headcount: null,
      region: null,
    });

    expect(updated.values).toEqual({ headcount: null, region: null });
    database.close();
  });

  it("matches the table name case-insensitively and the Record ID exactly", async () => {
    const database = await seeded();

    expect(
      database.updateRecord("customer", "CUST-0001", { name: "Acme Group" })
        .recordId,
    ).toBe("CUST-0001");
    expect(
      codeOf(() => database.updateRecord("Customer", "cust-0001", {})),
    ).toBe("DB_RECORD_NOT_FOUND");
    database.close();
  });

  it("checks the record exists even when no columns are given", async () => {
    const database = await seeded();

    expect(database.updateRecord("Customer", "CUST-0001", {})).toEqual({
      recordId: "CUST-0001",
      values: {},
    });
    expect(
      codeOf(() => database.updateRecord("Customer", "CUST-0002", {})),
    ).toBe("DB_RECORD_NOT_FOUND");
    database.close();
  });

  it("survives a save and reload, so an edit is in the saved bytes", async () => {
    const database = await seeded();
    database.updateRecord("Customer", "CUST-0001", { name: "Acme Holdings" });
    const bytes = database.serialize();
    database.close();

    const reopened = await Database.open(bytes);
    expect(reopened.readRecords("Customer")[0]).toMatchObject({
      record_id: "CUST-0001",
      name: "Acme Holdings",
    });
    reopened.close();
  });

  it("keeps editing after a save, because serialize resets the connection", async () => {
    const database = await seeded();
    database.serialize();

    // A foreign key violation after a save proves enforcement is still on.
    expect(
      codeOf(() =>
        database.updateRecord("Customer", "CUST-0001", { region: "REG-9999" }),
      ),
    ).toBe("DB_FOREIGN_KEY_VIOLATION");
    database.close();
  });
});

describe("updateRecord refusals", () => {
  it("refuses an unknown table", async () => {
    const database = await seeded();
    expect(
      codeOf(() =>
        database.updateRecord("Supplier", "SUP-0001", { name: "x" }),
      ),
    ).toBe("DB_TABLE_NOT_FOUND");
    database.close();
  });

  it("refuses an unknown record", async () => {
    const database = await seeded();
    expect(
      codeOf(() =>
        database.updateRecord("Customer", "CUST-0404", { name: "Globex" }),
      ),
    ).toBe("DB_RECORD_NOT_FOUND");
    database.close();
  });

  it("refuses an unknown column", async () => {
    const database = await seeded();
    expect(
      codeOf(() =>
        database.updateRecord("Customer", "CUST-0001", { founded: 1999 }),
      ),
    ).toBe("DB_UNKNOWN_COLUMN");
    database.close();
  });

  it("refuses two keys that resolve to the same column", async () => {
    const database = await seeded();
    expect(
      codeOf(() =>
        database.updateRecord("Customer", "CUST-0001", {
          name: "Acme",
          NAME: "Globex",
        }),
      ),
    ).toBe("DB_DUPLICATE_UPDATE_COLUMN");
    database.close();
  });

  it("refuses any attempt to change the Record ID, whatever its casing", async () => {
    const database = await seeded();

    expect(
      codeOf(() =>
        database.updateRecord("Customer", "CUST-0001", {
          record_id: "CUST-9999",
        }),
      ),
    ).toBe("DB_RECORD_ID_IMMUTABLE");
    expect(
      codeOf(() =>
        database.updateRecord("Customer", "CUST-0001", {
          RECORD_ID: "CUST-9999",
        }),
      ),
    ).toBe("DB_RECORD_ID_IMMUTABLE");
    // Even setting it to the value it already holds is refused, so there is no
    // path through this method that writes the column at all.
    expect(
      codeOf(() =>
        database.updateRecord("Customer", "CUST-0001", {
          record_id: "CUST-0001",
        }),
      ),
    ).toBe("DB_RECORD_ID_IMMUTABLE");
    database.close();
  });

  it("refuses a value that does not fit the column type, and writes nothing", async () => {
    const database = await seeded();

    expect(
      codeOf(() =>
        database.updateRecord("Customer", "CUST-0001", {
          name: "Acme Holdings",
          headcount: "twelve",
        }),
      ),
    ).toBe("DB_INVALID_NUMBER");
    expect(
      codeOf(() =>
        database.updateRecord("Customer", "CUST-0001", { headcount: 1.5 }),
      ),
    ).toBe("DB_INVALID_NUMBER");
    expect(
      codeOf(() =>
        database.updateRecord("Customer", "CUST-0001", { active: "maybe" }),
      ),
    ).toBe("DB_INVALID_BOOLEAN");

    // Conversion happens before any statement runs, so the valid column in the
    // first refusal above was not written either.
    expect(database.readRecords("Customer")[0]).toEqual({
      record_id: "CUST-0001",
      name: "Acme",
      region: "REG-0001",
      headcount: 12,
      active: true,
    });
    database.close();
  });

  it("names the table and column in a conversion refusal", async () => {
    const database = await seeded();
    let message = "";
    try {
      database.updateRecord("Customer", "CUST-0001", { headcount: "twelve" });
    } catch (error) {
      message = error instanceof Error ? error.message : "";
    }
    expect(message).toContain('table "Customer"');
    expect(message).toContain('column "headcount"');
    database.close();
  });

  it("refuses a foreign key that points at no record", async () => {
    const database = await seeded();
    expect(
      codeOf(() =>
        database.updateRecord("Customer", "CUST-0001", { region: "REG-9999" }),
      ),
    ).toBe("DB_FOREIGN_KEY_VIOLATION");
    expect(database.readRecords("Customer")[0]).toMatchObject({
      region: "REG-0001",
    });
    database.close();
  });

  it("accepts a foreign key that points at an existing record", async () => {
    const database = await seeded();
    expect(
      database.updateRecord("Customer", "CUST-0001", { region: "REG-0002" })
        .values,
    ).toEqual({ region: "REG-0002" });
    database.close();
  });

  it("refuses emptying a column declared non-nullable", async () => {
    const database = await seeded();
    expect(
      codeOf(() =>
        database.updateRecord("Customer", "CUST-0001", { name: null }),
      ),
    ).toBe("DB_NOT_NULL_VIOLATION");
    database.close();
  });

  it("cannot be talked past by a raw statement, because the trigger holds", async () => {
    const database = await seeded();
    // The library refusal above is the friendly path; storage enforcement is
    // what makes Record ID immutability true rather than merely encouraged.
    expect(() =>
      database.sql.run(
        `UPDATE ${quoteIdentifier("Customer")} SET ${quoteIdentifier("record_id")} = ? WHERE ${quoteIdentifier("record_id")} = ?;`,
        ["CUST-9999", "CUST-0001"],
      ),
    ).toThrow(/immutable/iu);
    database.close();
  });
});
