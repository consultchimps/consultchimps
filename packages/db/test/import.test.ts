import { isConsultChimpsError } from "@consultchimps/core";
import type { Table, TableRow } from "@consultchimps/tabular";
import { describe, expect, it } from "vitest";

import {
  Database,
  databaseTableToTable,
  importTable,
  importTables,
  inferColumnTypes,
  parseCsvTable,
  suggestRecordIdPrefix,
  suggestTableName,
} from "../src/index.js";

/** A one-column table, so an inference rule can be stated as its values. */
function column(values: Array<string | number | boolean | null>): Table {
  return {
    columns: ["Value"],
    rows: values.map((value) => ({ Value: value })),
  };
}

function typeOfColumn(
  values: Array<string | number | boolean | null>,
): string | undefined {
  return inferColumnTypes(column(values))[0]?.type;
}

async function codeOf(run: () => unknown): Promise<string | undefined> {
  try {
    await run();
  } catch (error) {
    return isConsultChimpsError(error) ? error.code : undefined;
  }
  return undefined;
}

const CUSTOMERS: Table = {
  columns: ["Customer", "Region", "Active", "Score", "Opened"],
  rows: [
    {
      Customer: "Acme",
      Region: "North",
      Active: "true",
      Score: "12.5",
      Opened: "2026-01-31",
    },
    {
      Customer: "Beta",
      Region: "South",
      Active: "false",
      Score: "8",
      Opened: "2026-02-01",
    },
  ],
};

describe("inferColumnTypes", () => {
  it("takes a type only when every value that is not blank fits it", () => {
    expect(typeOfColumn(["1", "2", null, "  "])).toBe("integer");
    expect(typeOfColumn(["1", "2.5"])).toBe("real");
    expect(typeOfColumn(["true", "NO", "Yes"])).toBe("boolean");
    expect(typeOfColumn([true, false])).toBe("boolean");
    expect(typeOfColumn(["2026-01-31", "2026-12-01"])).toBe("date");
    expect(typeOfColumn(["2026-01-31T09:30:00.000Z"])).toBe("date");
  });

  it("falls back to text for a mixed column and an empty one", () => {
    expect(typeOfColumn(["1", "yes"])).toBe("text");
    expect(typeOfColumn(["2026-01-31", "01/02/2026"])).toBe("text");
    expect(typeOfColumn([])).toBe("text");
    expect(typeOfColumn([null, "", "   "])).toBe("text");
  });

  it("keeps a value text whenever reading it as a number would change it", () => {
    // A padded reference code, an exponent form, a grouped number, a leading
    // plus, and a whole number past the exactly representable range.
    expect(typeOfColumn(["007", "008"])).toBe("text");
    expect(typeOfColumn(["1e5"])).toBe("text");
    expect(typeOfColumn(["1,000"])).toBe("text");
    expect(typeOfColumn(["+5"])).toBe("text");
    expect(typeOfColumn(["9007199254740993"])).toBe("text");
    expect(typeOfColumn(["0", "-3", "0.5"])).toBe("real");
  });

  it("refuses a date that is not on the calendar, and an out-of-range time", () => {
    expect(typeOfColumn(["2026-02-30"])).toBe("text");
    expect(typeOfColumn(["2026-13-01"])).toBe("text");
    expect(typeOfColumn(["2024-02-29"])).toBe("date");
    expect(typeOfColumn(["2026-01-31T25:00:00Z"])).toBe("text");
  });

  it("reports how many values decided each column", () => {
    expect(inferColumnTypes(column(["1", null, "2"]))[0]?.valueCount).toBe(2);
  });
});

describe("suggestTableName and suggestRecordIdPrefix", () => {
  it("suggests a safe identifier and a readable prefix", () => {
    expect(suggestTableName("Customer List 2026")).toBe("Customer_List_2026");
    expect(suggestTableName("  ")).toBe("table");
    expect(suggestTableName("sqlite_stat")).toBe("table_sqlite_stat");
    expect(suggestTableName("__proto__")).toBe("proto");
    expect(suggestRecordIdPrefix("Customers")).toBe("CUST");
    expect(suggestRecordIdPrefix("Sales Orders")).toBe("SO");
    expect(suggestRecordIdPrefix("__")).toBe("REC");
  });
});

describe("importTable", () => {
  it("creates the table, infers its columns, and generates Record IDs", async () => {
    const database = await Database.create();

    const imported = importTable(database, CUSTOMERS, {
      name: " Customers ",
      recordId: { prefix: " CUST ", padding: 4 },
    });

    expect(imported.name).toBe("Customers");
    expect(imported.recordId).toEqual({ prefix: "CUST", padding: 4 });
    expect(imported.rowCount).toBe(2);
    expect(imported.firstRecordId).toBe("CUST-0001");
    expect(imported.lastRecordId).toBe("CUST-0002");
    expect(imported.columns.map((entry) => [entry.name, entry.type])).toEqual([
      ["Customer", "text"],
      ["Region", "text"],
      ["Active", "boolean"],
      ["Score", "real"],
      ["Opened", "date"],
    ]);

    expect(databaseTableToTable(database, "Customers").rows).toEqual([
      {
        record_id: "CUST-0001",
        Customer: "Acme",
        Region: "North",
        Active: true,
        Score: 12.5,
        Opened: "2026-01-31",
      },
      {
        record_id: "CUST-0002",
        Customer: "Beta",
        Region: "South",
        Active: false,
        Score: 8,
        Opened: "2026-02-01",
      },
    ]);
    database.close();
  });

  it("creates an empty table when the source has a header row and no rows", async () => {
    const database = await Database.create();

    const imported = importTable(
      database,
      { columns: ["Customer"], rows: [] },
      { name: "Customers", recordId: { prefix: "CUST", padding: 4 } },
    );

    expect(imported.rowCount).toBe(0);
    expect(imported.firstRecordId).toBeNull();
    expect(database.countRecords("Customers")).toBe(0);
    database.close();
  });

  it("imports a table read back out of a workspace, without its Record ID", async () => {
    const database = await Database.create();
    importTable(database, CUSTOMERS, {
      name: "Customers",
      recordId: { prefix: "CUST", padding: 4 },
    });

    const copied = importTable(
      database,
      databaseTableToTable(database, "Customers"),
      { name: "CustomerCopy", recordId: { prefix: "COPY", padding: 4 } },
    );

    expect(copied.ignoredColumns).toEqual(["record_id"]);
    expect(copied.columns.map((entry) => entry.name)).not.toContain(
      "record_id",
    );
    expect(copied.firstRecordId).toBe("COPY-0001");
    database.close();
  });

  it("refuses a name already taken, whatever its case, before writing", async () => {
    const database = await Database.create();
    importTable(database, CUSTOMERS, {
      name: "Customers",
      recordId: { prefix: "CUST", padding: 4 },
    });

    expect(
      await codeOf(() =>
        importTable(database, CUSTOMERS, {
          name: "customers",
          recordId: { prefix: "OTHR", padding: 4 },
        }),
      ),
    ).toBe("DB_TABLE_EXISTS");
    expect(database.getSchema()).toHaveLength(1);
    database.close();
  });

  it("refuses an empty prefix, an out-of-range padding, and a reserved name", async () => {
    const database = await Database.create();
    const options = {
      name: "Customers",
      recordId: { prefix: "CUST", padding: 4 },
    };

    expect(
      await codeOf(() =>
        importTable(database, CUSTOMERS, {
          ...options,
          recordId: { prefix: "   ", padding: 4 },
        }),
      ),
    ).toBe("DB_INVALID_RECORD_ID_CONFIG");
    expect(
      await codeOf(() =>
        importTable(database, CUSTOMERS, {
          ...options,
          recordId: { prefix: "CUST", padding: 500 },
        }),
      ),
    ).toBe("DB_INVALID_RECORD_ID_CONFIG");
    expect(
      await codeOf(() =>
        importTable(database, CUSTOMERS, {
          ...options,
          name: "_consultchimps_notes",
        }),
      ),
    ).toBe("DB_RESERVED_IDENTIFIER");
    expect(
      await codeOf(() =>
        importTable(
          database,
          { columns: ["_consultchimps_id"], rows: [] },
          options,
        ),
      ),
    ).toBe("DB_RESERVED_IDENTIFIER");
    expect(
      await codeOf(() =>
        importTable(database, { columns: [], rows: [] }, options),
      ),
    ).toBe("DB_IMPORT_NO_COLUMNS");
    expect(database.getSchema()).toEqual([]);
    database.close();
  });
});

describe("importTables", () => {
  it("creates every chosen table in one call", async () => {
    const database = await Database.create();

    const imported = importTables(database, [
      {
        name: "Customers",
        recordId: { prefix: "CUST", padding: 4 },
        table: CUSTOMERS,
      },
      {
        name: "Regions",
        recordId: { prefix: "REG", padding: 3 },
        table: { columns: ["Region"], rows: [{ Region: "North" }] },
      },
    ]);

    expect(imported.map((entry) => entry.name)).toEqual([
      "Customers",
      "Regions",
    ]);
    expect(database.countRecords("Regions")).toBe(1);
    expect(
      databaseTableToTable(database, "Regions").rows[0]?.["record_id"],
    ).toBe("REG-001");
    database.close();
  });

  it("refuses two requests that would create the same table", async () => {
    const database = await Database.create();

    expect(
      await codeOf(() =>
        importTables(database, [
          {
            name: "Customers",
            recordId: { prefix: "CUST", padding: 4 },
            table: CUSTOMERS,
          },
          {
            name: "customers",
            recordId: { prefix: "OTHR", padding: 4 },
            table: CUSTOMERS,
          },
        ]),
      ),
    ).toBe("DB_TABLE_EXISTS");
    expect(database.getSchema()).toEqual([]);
    database.close();
  });

  it("refuses an empty request list", async () => {
    const database = await Database.create();
    expect(await codeOf(() => importTables(database, []))).toBe(
      "DB_IMPORT_NOTHING_SELECTED",
    );
    database.close();
  });

  it("leaves the workspace untouched when a later table fails mid-import", async () => {
    const database = await Database.create();
    importTable(database, CUSTOMERS, {
      name: "Existing",
      recordId: { prefix: "EX", padding: 4 },
    });

    // A source that stops being readable partway through the import: the value
    // is read once while the columns are inferred and again while the row is
    // inserted, and the second read fails. That is the shape of every failure
    // validation cannot catch up front (a worker that dies, a file that goes
    // away), and it is the one an import has to survive without leaving half a
    // workbook behind.
    let reads = 0;
    const failingRow = Object.defineProperty({}, "Value", {
      enumerable: true,
      get: (): string => {
        reads += 1;
        if (reads > 1) {
          throw new Error("The file could not be read to the end.");
        }
        return "2";
      },
    }) as TableRow;

    await expect(async () =>
      importTables(database, [
        {
          name: "Fine",
          recordId: { prefix: "FIN", padding: 4 },
          table: CUSTOMERS,
        },
        {
          name: "Broken",
          recordId: { prefix: "BRK", padding: 4 },
          table: { columns: ["Value"], rows: [failingRow] },
        },
      ]),
    ).rejects.toThrow("could not be read");

    // Neither the table that had already been created nor the one that failed
    // survives, and the workspace still holds only what it held before.
    expect(database.getSchema().map((schema) => schema.name)).toEqual([
      "Existing",
    ]);
    database.close();
  });

  it("survives a save and a reopen with its schema, rows, and next id", async () => {
    const database = await Database.create();
    importTables(database, [
      {
        name: "Customers",
        recordId: { prefix: "CUST", padding: 4 },
        table: parseCsvTable(
          "Customer,Region,Score\nAcme,North,10\nBeta,South,20\n",
        ),
      },
    ]);
    const bytes = database.serialize();
    database.close();

    const reopened = await Database.open(bytes);
    const schema = reopened.getSchema();
    expect(schema).toHaveLength(1);
    expect(schema[0]?.recordId).toEqual({ prefix: "CUST", padding: 4 });
    expect(schema[0]?.columns.map((entry) => entry.type)).toEqual([
      "text",
      "text",
      "integer",
    ]);
    expect(reopened.countRecords("Customers")).toBe(2);
    // The counter carried over, so a record added after reopening continues the
    // sequence rather than restarting it.
    expect(
      reopened.insertRecord("Customers", { Customer: "Gamma" }).recordId,
    ).toBe("CUST-0003");
    reopened.close();
  });
});
