import {
  isConsultChimpsError,
  type ConsultChimpsError,
} from "@consultchimps/core";
import type { Table, TableRow } from "@consultchimps/tabular";
import { describe, expect, it } from "vitest";

import {
  assertSafeIdentifier,
  Database,
  databaseTableToTable,
  importTable,
  importTables,
  inferColumnTypes,
  isValueConversionError,
  parseCsvTable,
  suggestRecordIdPrefix,
  suggestTableName,
  truncateIdentifier,
  importedTableSchema,
  type TableSchema,
  MAX_IDENTIFIER_LENGTH,
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
    // A padded reference code, an exponent form, a grouped number, and a
    // leading plus: each says something the number would not say back.
    expect(typeOfColumn(["007", "008"])).toBe("text");
    expect(typeOfColumn(["1e5"])).toBe("text");
    expect(typeOfColumn(["1,000"])).toBe("text");
    expect(typeOfColumn(["+5"])).toBe("text");
    expect(typeOfColumn(["0", "-3", "0.5"])).toBe("real");
  });

  it("holds both numeric paths to the same round trip", () => {
    // Whole numbers: past the exactly representable range, and past the range
    // an integer column stores, both of which read back as another number.
    expect(typeOfColumn(["9007199254740993"])).toBe("text");
    expect(typeOfColumn(["9007199254740992"])).toBe("text");
    expect(typeOfColumn(["9007199254740991"])).toBe("integer");
    expect(typeOfColumn([`1${"0".repeat(400)}`])).toBe("text");

    // Decimals: more precision than a number can hold, underflow to zero, and
    // overflow to Infinity. Before the shared predicate the first of these was
    // accepted, because the decimal path only asked whether Number() was
    // finite.
    expect(typeOfColumn(["0.12345678901234567890"])).toBe("text");
    expect(typeOfColumn([`0.${"0".repeat(400)}1`])).toBe("text");
    expect(typeOfColumn([`1${"0".repeat(400)}.5`])).toBe("text");

    // A decimal that does survive the conversion, including one written with a
    // trailing zero, which is the same value spelled differently.
    expect(typeOfColumn(["12.5", "0.25"])).toBe("real");
    expect(typeOfColumn(["12.50"])).toBe("real");
    expect(typeOfColumn(["0.1"])).toBe("real");
  });

  it("stores a number column without changing any of its values", async () => {
    const database = await Database.create();
    importTable(
      database,
      {
        columns: ["Score"],
        rows: [{ Score: "12.50" }, { Score: "0.25" }],
      },
      { name: "Scores", recordId: { prefix: "SC", padding: 4 } },
    );
    expect(
      databaseTableToTable(database, "Scores").rows.map((row) => row["Score"]),
    ).toEqual([12.5, 0.25]);
    database.close();
  });

  it("refuses a date that is not on the calendar", () => {
    expect(typeOfColumn(["2026-02-30"])).toBe("text");
    expect(typeOfColumn(["2026-13-01"])).toBe("text");
    expect(typeOfColumn(["2024-02-29"])).toBe("date");
    expect(typeOfColumn(["2023-02-29"])).toBe("text");
  });

  it("judges a timestamp as a whole, not part by part", () => {
    // Hour 24 is the end of a day and nothing else.
    expect(typeOfColumn(["2026-01-31T24:00:00Z"])).toBe("date");
    expect(typeOfColumn(["2026-01-31T24:00"])).toBe("date");
    expect(typeOfColumn(["2026-01-31T24:30:00Z"])).toBe("text");
    expect(typeOfColumn(["2026-01-31T24:00:01Z"])).toBe("text");
    expect(typeOfColumn(["2026-01-31T24:00:00.500Z"])).toBe("text");
    expect(typeOfColumn(["2026-01-31T24:00:00.000Z"])).toBe("date");
    expect(typeOfColumn(["2026-01-31T25:00:00Z"])).toBe("text");

    // A leap second sits at 23:59:60 UTC and nowhere else, so the offset is
    // part of the rule rather than a separate check that happens to pass.
    expect(typeOfColumn(["2026-12-31T23:59:60Z"])).toBe("date");
    expect(typeOfColumn(["2026-12-31T23:59:60+00:00"])).toBe("date");
    expect(typeOfColumn(["2026-12-31T23:59:60+05:00"])).toBe("text");
    expect(typeOfColumn(["2026-12-31T23:59:60-00:00"])).toBe("text");
    expect(typeOfColumn(["2026-12-31T23:59:60"])).toBe("text");
    expect(typeOfColumn(["2026-01-31T09:30:60Z"])).toBe("text");
    expect(typeOfColumn(["2026-01-31T23:58:60Z"])).toBe("text");
    expect(typeOfColumn(["2026-01-31T09:61:00Z"])).toBe("text");

    // An offset is a real offset.
    expect(typeOfColumn(["2026-01-31T09:30:00+05:30"])).toBe("date");
    expect(typeOfColumn(["2026-01-31T09:30:00-08:00"])).toBe("date");
    expect(typeOfColumn(["2026-01-31T09:30:00+99:99"])).toBe("text");
    expect(typeOfColumn(["2026-01-31T09:30:00+24:00"])).toBe("text");
    expect(typeOfColumn(["2026-01-31T09:30:00+05:60"])).toBe("text");
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

describe("dates a date column will hold", () => {
  it("stores what inference judged, without the spaces around it", async () => {
    const database = await Database.create();

    const imported = importTable(
      database,
      {
        columns: ["Opened"],
        rows: [{ Opened: " 2026-01-31 " }, { Opened: "2026-02-01" }],
      },
      { name: "Events", recordId: { prefix: "EV", padding: 4 } },
    );

    // Inference judged the trimmed text, so the trimmed text is what is stored.
    expect(imported.columns[0]?.type).toBe("date");
    expect(
      databaseTableToTable(database, "Events").rows.map((row) => row["Opened"]),
    ).toEqual(["2026-01-31", "2026-02-01"]);
    database.close();
  });

  it("reads a cell of nothing but spaces as no value at all", async () => {
    const database = await Database.create();

    importTable(
      database,
      {
        columns: ["Opened", "Note"],
        rows: [
          { Opened: "   ", Note: "  " },
          { Opened: "2026-01-31", Note: "x" },
        ],
      },
      { name: "Events", recordId: { prefix: "EV", padding: 4 } },
    );

    // Null rather than whitespace, in a text column as well as a typed one, so
    // an empty cell is findable as empty.
    expect(databaseTableToTable(database, "Events").rows[0]).toMatchObject({
      Opened: null,
      Note: null,
    });
    expect(
      database.sql.selectValue(
        'SELECT count(*) FROM "Events" WHERE "Note" IS NULL;',
      ),
    ).toBe(1);
    database.close();
  });

  it("refuses a value that is not a date, whatever writes it", async () => {
    const database = await Database.create();
    database.createTable({
      name: "Events",
      columns: [{ name: "Opened", type: "date" }],
      foreignKeys: [],
      recordId: { prefix: "EV", padding: 4 },
    });

    // The single conversion point every write goes through, so the record grid
    // inherits this without having to remember it.
    expect(
      await codeOf(() => database.insertRecord("Events", { Opened: "hello" })),
    ).toBe("DB_INVALID_DATE");
    expect(
      await codeOf(() =>
        database.insertRecord("Events", { Opened: "01/02/2026" }),
      ),
    ).toBe("DB_INVALID_DATE");
    expect(
      await codeOf(() =>
        database.insertRecord("Events", { Opened: "2026-02-30" }),
      ),
    ).toBe("DB_INVALID_DATE");

    // A real date, written with a spreadsheet's padding, is stored as the date.
    database.insertRecord("Events", { Opened: " 2026-01-31T09:30:00Z " });
    expect(database.readRecords("Events")[0]?.["Opened"]).toBe(
      "2026-01-31T09:30:00Z",
    );
    // And a blank is an unset cell, as it is for every other typed column.
    database.insertRecord("Events", { Opened: "  " });
    expect(database.readRecords("Events")[1]?.["Opened"]).toBeNull();
    database.close();
  });

  it("stores a padded number and boolean as the value inference judged", async () => {
    // The same question as the date, asked of the other typed columns: the
    // conversion trims before reading them, so what inference judged is what is
    // stored, and nothing keeps the padding.
    const database = await Database.create();

    importTable(
      database,
      {
        columns: ["Score", "Active"],
        rows: [{ Score: " 12.5 ", Active: " TRUE " }],
      },
      { name: "Rows", recordId: { prefix: "R", padding: 4 } },
    );

    expect(databaseTableToTable(database, "Rows").rows[0]).toMatchObject({
      Score: 12.5,
      Active: true,
    });
    // And a grouped number is still text, so nothing about it is normalized.
    expect(typeOfColumn(["1,000"])).toBe("text");
    database.close();
  });

  it("reports a stored value that is not a date as damage", async () => {
    const database = await Database.create();
    database.createTable({
      name: "Events",
      columns: [{ name: "Opened", type: "date" }],
      foreignKeys: [],
      recordId: { prefix: "EV", padding: 4 },
    });
    // Written past the conversion, the way an external edit would.
    database.sql.run(
      'INSERT INTO "Events" (record_id, "Opened") VALUES (?, ?);',
      ["EV-9999", "not a date"],
    );

    expect(await codeOf(() => database.readRecords("Events"))).toBe(
      "DB_CORRUPT_STORED_VALUE",
    );
    database.close();
  });
});

describe("a value a column will not take", () => {
  /**
   * The one expectation all three conversions are held to, so a caller is never
   * told which type complained without being told where. Written once on
   * purpose: three copies would let one of them fall behind.
   */
  async function expectsTableAndColumn(
    type: "boolean" | "integer" | "date",
    value: string,
    code: string,
  ): Promise<void> {
    const database = await Database.create();
    database.createTable({
      name: "Customer",
      columns: [{ name: "Field", type }],
      foreignKeys: [],
      recordId: { prefix: "CUST", padding: 4 },
    });

    let caught: unknown;
    try {
      database.insertRecord("Customer", { Field: value });
    } catch (error) {
      caught = error;
    }

    expect(isConsultChimpsError(caught)).toBe(true);
    const failure = caught as ConsultChimpsError;
    expect(failure.code).toBe(code);
    // Named where a person would look for it, in the sentence and in the
    // machine-readable details alike.
    expect(failure.message).toContain('table "Customer"');
    expect(failure.message).toContain('column "Field"');
    expect(failure.details).toMatchObject({
      column: "Field",
      table: "Customer",
    });
    // And never the offending value, which is imported cell content.
    expect(failure.message).not.toContain(value);
    database.close();
  }

  it("names the table and the column for a date", async () => {
    await expectsTableAndColumn("date", "01/02/2026", "DB_INVALID_DATE");
  });

  it("names the table and the column for a boolean", async () => {
    await expectsTableAndColumn("boolean", "perhaps", "DB_INVALID_BOOLEAN");
  });

  it("names the table and the column for a number", async () => {
    await expectsTableAndColumn("integer", "twelve", "DB_INVALID_NUMBER");
  });

  it("names them for a whole number outside the range stored exactly", async () => {
    await expectsTableAndColumn(
      "integer",
      "9007199254740993",
      "DB_INVALID_NUMBER",
    );
  });

  it("recognises a conversion failure by its marker, not by its code", async () => {
    // The wrapper above asks this one question. A conversion added later
    // inherits the context by raising the same marker, rather than by being
    // added to a list somewhere else.
    const database = await Database.create();
    database.createTable({
      name: "Customer",
      columns: [{ name: "Field", type: "date" }],
      foreignKeys: [],
      recordId: { prefix: "CUST", padding: 4 },
    });
    let caught: unknown;
    try {
      database.insertRecord("Customer", { Field: "nope" });
    } catch (error) {
      caught = error;
    }
    expect(isValueConversionError(caught)).toBe(true);

    // A failure that is not about a value is not marked, so the wrapper leaves
    // it alone rather than dressing it up with a column it has nothing to do
    // with.
    let other: unknown;
    try {
      database.insertRecord("Customer", { Missing: "x" });
    } catch (error) {
      other = error;
    }
    expect(isConsultChimpsError(other)).toBe(true);
    expect((other as ConsultChimpsError).code).toBe("DB_UNKNOWN_COLUMN");
    expect(isValueConversionError(other)).toBe(false);
    database.close();
  });
});

describe("the schema a preview promises", () => {
  const OPTIONS = {
    name: "Preview",
    recordId: { prefix: "PV", padding: 4 },
  };

  /** The table an import creates, read back from the database it created it in. */
  async function createdSchema(table: Table): Promise<TableSchema> {
    const database = await Database.create();
    importTable(database, table, OPTIONS);
    const [schema] = database.getSchema();
    database.close();
    return schema as TableSchema;
  }

  const CASES: ReadonlyArray<readonly [string, Table]> = [
    [
      "a header past the limit",
      {
        columns: ["z".repeat(200), `${"z".repeat(200)}_2`],
        rows: [{ [`${"z".repeat(200)}_2`]: "v" }],
      },
    ],
    [
      "headers that collide once shortened",
      {
        columns: [`${"y".repeat(199)}A`, `${"y".repeat(199)}B`],
        rows: [{ [`${"y".repeat(199)}A`]: "v" }],
      },
    ],
    [
      "a column holding nothing but spaces",
      { columns: ["Note", "Score"], rows: [{ Note: "   ", Score: " 12 " }] },
    ],
  ];

  it.each(CASES)(
    "promises the schema the import builds, for %s",
    async (_label, table) => {
      const preview = importedTableSchema(table, OPTIONS);
      // Deeply equal to what the database ended up holding, not merely similar.
      expect(preview).toEqual(await createdSchema(table));
    },
  );

  it.each(CASES)(
    "never previews a name the schema refuses, for %s",
    (_label, table) => {
      for (const column of importedTableSchema(table, OPTIONS).columns) {
        expect(() => assertSafeIdentifier(column.name, "column")).not.toThrow();
      }
    },
  );
});

describe("column names that have to fit", () => {
  const WIDE = "\u{10400}";

  /**
   * Import one header row and report the names the table ended up with.
   *
   * The headers are the ones a reader hands over, which are already distinct:
   * both the delimited-text and worksheet readers number a repeated header
   * before the import sees it. That is where the too-long name comes from in
   * the first place, since numbering a header already at the limit is what
   * pushes it past one.
   */
  async function columnsFor(headers: string[]): Promise<{
    names: string[];
    renamed: ReadonlyArray<{ from: string; to: string }>;
    rows: TableRow[];
  }> {
    expect(new Set(headers).size).toBe(headers.length);
    const database = await Database.create();
    const row: TableRow = Object.create(null);
    headers.forEach((header, index) => {
      row[header] = `v${index}`;
    });
    const imported = importTable(
      database,
      { columns: headers, rows: [row] },
      { name: "Wide", recordId: { prefix: "W", padding: 4 } },
    );
    const stored = databaseTableToTable(database, "Wide");
    database.close();
    return {
      names: imported.columns.map((column) => column.name),
      renamed: imported.renamedColumns,
      rows: stored.rows,
    };
  }

  /** Every name a table was given is one the schema will take. */
  function expectStorable(names: string[]): void {
    expect(new Set(names).size).toBe(names.length);
    for (const name of names) {
      expect(name.length).toBeLessThanOrEqual(MAX_IDENTIFIER_LENGTH);
      expect(name.isWellFormed()).toBe(true);
      expect(() => assertSafeIdentifier(name, "column")).not.toThrow();
    }
  }

  it("takes a header the reader already numbered past the limit", async () => {
    // Two copies of a 200-character header reach the import as the 200 and the
    // 202 the reader made of them. Checking the limit after the numbering, in
    // another layer, refused the whole import over a column name nobody can
    // edit, for a duplicate the import was meant to number for them.
    const header = "a".repeat(200);
    const { names, renamed, rows } = await columnsFor([header, `${header}_2`]);

    expectStorable(names);
    expect(names).toHaveLength(2);
    // Both values are still there, under the names the result reports.
    expect(rows[0]?.[names[0] as string]).toBe("v0");
    expect(rows[0]?.[names[1] as string]).toBe("v1");
    expect(renamed.map((entry) => entry.from)).toEqual([header, `${header}_2`]);
  });

  it("numbers headers that differ only past the cut, keeping every value", async () => {
    const base = "a".repeat(199);
    const { names, rows } = await columnsFor([`${base}X`, `${base}Y`]);

    expectStorable(names);
    expect(rows[0]?.[names[0] as string]).toBe("v0");
    expect(rows[0]?.[names[1] as string]).toBe("v1");
  });

  it("cuts a header between whole characters when the limit falls inside one", async () => {
    // The budget lands exactly where the wide character starts, so a cut by
    // code unit would leave half of it behind.
    const budget = MAX_IDENTIFIER_LENGTH - "_6".length;
    const header = "a".repeat(budget - 1) + WIDE;
    const { names } = await columnsFor([header, `${header}_2`]);

    expectStorable(names);
  });

  it("keeps twelve long headers inside the limit, with two-digit numbers", async () => {
    const header = "b".repeat(200);
    const { names } = await columnsFor([
      header,
      ...Array.from(
        { length: 11 },
        (_unused, index) => `${header}_${index + 2}`,
      ),
    ]);

    expect(names).toHaveLength(12);
    expectStorable(names);
  });

  it("never produces a column name the schema would refuse", async () => {
    // A stand-in for the shapes a real header row takes: at the limit, past it,
    // differing only past the cut, and a wide character on the boundary. The
    // cases are enumerated rather than drawn at random, so a failure here fails
    // the same way twice.
    const lengths = [1, 2, 190, 197, 198, 199, 200, 201, 250, 300];
    for (const length of lengths) {
      for (const filler of ["c", WIDE]) {
        const base = filler.repeat(length);
        const { names } = await columnsFor([
          base,
          `${base}x`,
          `${base}_2`,
          `${base}${WIDE}`,
        ]);
        expect(names).toHaveLength(4);
        expectStorable(names);
      }
    }
  });

  it("leaves a unique header at the limit exactly as it was", async () => {
    // Nothing will be numbered, so nothing has to make room for a number.
    // Taking that room anyway shortened a schema the person cannot rename.
    const header = "d".repeat(200);
    const { names, renamed } = await columnsFor([header]);

    expect(names).toEqual([header]);
    expect(renamed).toEqual([]);
    expectStorable(names);
  });

  it("shortens the headers that fold together and leaves the rest", async () => {
    const unique = "e".repeat(200);
    const pair = "f".repeat(200);
    const { names, renamed } = await columnsFor([unique, pair, `${pair}_2`]);

    // The one that stands alone keeps every character it had.
    expect(names[0]).toBe(unique);
    expect(renamed.some((entry) => entry.from === unique)).toBe(false);
    // The two that fold together are numbered, inside the limit.
    expect(names[1]).not.toBe(names[2]);
    expectStorable(names);
  });

  it("leaves a header that already fits exactly as it was", async () => {
    const { names, renamed } = await columnsFor(["Customer", "Region"]);
    expect(names).toEqual(["Customer", "Region"]);
    expect(renamed).toEqual([]);
  });
});

describe("identifiers made of whole characters", () => {
  // U+10400 DESERET CAPITAL LONG I: one character, two UTF-16 code units, which
  // is what a slice at a code-unit index can cut in half.
  const WIDE = "\u{10400}";

  it("never cuts a character in half when it shortens a name", () => {
    // Long enough that the cut lands exactly where the wide character starts.
    const budget = MAX_IDENTIFIER_LENGTH - "table_".length;
    const source = "a".repeat(budget - 1) + WIDE;

    // The defect this pins down: cutting at the same budget by code units
    // leaves half a character behind, which is not text at all.
    expect(source.slice(0, budget).isWellFormed()).toBe(false);

    const suggested = suggestTableName(source);
    expect(suggested.isWellFormed()).toBe(true);
    expect(suggested.length).toBeLessThanOrEqual(budget);
    expect(suggested).toBe("a".repeat(budget - 1));
    expect(() => assertSafeIdentifier(suggested, "table")).not.toThrow();
  });

  it("shortens a name made only of wide characters to whole ones", () => {
    const suggested = suggestTableName(WIDE.repeat(300));

    expect(suggested.isWellFormed()).toBe(true);
    expect(suggested.length).toBeLessThanOrEqual(MAX_IDENTIFIER_LENGTH);
    expect(() => assertSafeIdentifier(suggested, "table")).not.toThrow();
  });

  it("truncates at the limit's own unit, whole characters at a time", () => {
    expect(truncateIdentifier("abcdef", 4)).toBe("abcd");
    expect(truncateIdentifier(`abc${WIDE}`, 4)).toBe("abc");
    expect(truncateIdentifier(`abc${WIDE}`, 5)).toBe(`abc${WIDE}`);
    expect(truncateIdentifier("abc", 10)).toBe("abc");
  });

  it("refuses an identifier that is not whole text", async () => {
    // Half a character survives every other check and is stored as U+FFFD, so
    // the name in the database would not be the name that was asked for.
    const half = "Custom\uD800er";
    expect(half.isWellFormed()).toBe(false);
    expect(await codeOf(() => assertSafeIdentifier(half, "table"))).toBe(
      "DB_INVALID_IDENTIFIER",
    );
    expect(
      await codeOf(() => assertSafeIdentifier("Region\uDC00", "column")),
    ).toBe("DB_INVALID_IDENTIFIER");
    // A whole character outside the basic plane is ordinary text.
    expect(() => assertSafeIdentifier(`Data${WIDE}`, "table")).not.toThrow();
  });

  it("refuses a column header that is not whole text, before writing", async () => {
    const database = await Database.create();
    expect(
      await codeOf(() =>
        importTable(
          database,
          { columns: ["Region\uD800"], rows: [] },
          { name: "Regions", recordId: { prefix: "REG", padding: 4 } },
        ),
      ),
    ).toBe("DB_INVALID_IDENTIFIER");
    expect(database.getSchema()).toEqual([]);
    database.close();
  });

  it("keeps a suggested Record ID prefix to whole characters too", () => {
    expect(suggestRecordIdPrefix(WIDE.repeat(8)).isWellFormed()).toBe(true);
    expect(suggestRecordIdPrefix(WIDE.repeat(8))).toBe(WIDE.repeat(4));
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

  it("creates nothing when a later table cannot be read", async () => {
    const database = await Database.create();
    importTable(database, CUSTOMERS, {
      name: "Existing",
      recordId: { prefix: "EX", padding: 4 },
    });

    // A source that cannot be read to the end. Every table is normalized,
    // judged, and checked before the first one is created, so a source that
    // fails takes the whole call down before anything exists, rather than
    // after some of it does. The creation and the rows still run inside one
    // transaction underneath, which covers what validation cannot foresee.
    const unreadableRow = Object.defineProperty({}, "Value", {
      enumerable: true,
      get: (): string => {
        throw new Error("The file could not be read to the end.");
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
          table: { columns: ["Value"], rows: [unreadableRow] },
        },
      ]),
    ).rejects.toThrow("could not be read");

    // Neither the table that would have come first nor the one that failed
    // exists, and the workspace still holds only what it held before.
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

describe("what a preview and an import agree about", () => {
  const OPTIONS = { name: "Empty", recordId: { prefix: "E", padding: 4 } };

  /** The code each half answers with, so the two can be compared as one. */
  async function codesFor(
    table: Table,
  ): Promise<{ preview: string | undefined; imported: string | undefined }> {
    const database = await Database.create();
    const preview = await codeOf(() => importedTableSchema(table, OPTIONS));
    const imported = await codeOf(() => importTable(database, table, OPTIONS));
    // Nothing was created whatever the answer was.
    expect(database.getSchema()).toEqual([]);
    database.close();
    return { preview, imported };
  }

  it("refuses a source with no columns the same way twice", async () => {
    const { preview, imported } = await codesFor({ columns: [], rows: [] });

    expect(preview).toBe("DB_IMPORT_NO_COLUMNS");
    expect(imported).toBe(preview);
  });

  it("refuses a source whose only column is the Record ID the same way", async () => {
    // The Record ID is generated, so a column called one contributes nothing:
    // the schema this would create has no columns of its own.
    const { preview, imported } = await codesFor({
      columns: ["record_id"],
      rows: [{ record_id: "CUST-0001" }],
    });

    expect(preview).toBe("DB_IMPORT_NO_COLUMNS");
    expect(imported).toBe(preview);
  });

  it("refuses a column listed twice under the same name, the same way", async () => {
    // A row is an object keyed by column name, so one property answers both
    // reads: importing this would store the first column's value twice.
    const row: TableRow = Object.create(null);
    row["Amount"] = 10;
    const { preview, imported } = await codesFor({
      columns: ["Amount", "Amount"],
      rows: [row],
    });

    expect(preview).toBe("DB_IMPORT_DUPLICATE_SOURCE_COLUMN");
    expect(imported).toBe(preview);
  });

  it("numbers two columns that differ only in case, and keeps both values", async () => {
    // These are two properties carrying two values, so nothing is ambiguous
    // and nothing is lost: the second is renamed, which is reported, and both
    // values are stored. Refusing it would refuse a source the row can answer.
    const row: TableRow = Object.create(null);
    row["Amount"] = 10;
    row["amount"] = 20;
    const database = await Database.create();
    const imported = importTable(
      database,
      { columns: ["Amount", "amount"], rows: [row] },
      { name: "Amounts", recordId: { prefix: "A", padding: 4 } },
    );
    const stored = databaseTableToTable(database, "Amounts");
    database.close();

    const names = imported.columns.map((entry) => entry.name);
    expect(new Set(names).size).toBe(2);
    expect(imported.renamedColumns).toHaveLength(1);
    expect(
      names
        .map((name) => stored.rows[0]?.[name])
        .sort((a, b) => Number(a) - Number(b)),
    ).toEqual([10, 20]);
  });

  it("refuses a table name the engine will not take, the same way", async () => {
    const database = await Database.create();
    const table: Table = { columns: ["Amount"], rows: [{ Amount: 1 }] };
    const options = { name: "   ", recordId: { prefix: "A", padding: 4 } };
    const preview = await codeOf(() => importedTableSchema(table, options));
    const imported = await codeOf(() => importTable(database, table, options));
    database.close();

    expect(preview).toBe("DB_INVALID_IDENTIFIER");
    expect(imported).toBe(preview);
  });

  it("refuses a Record ID configuration it cannot number with, the same way", async () => {
    const database = await Database.create();
    const table: Table = { columns: ["Amount"], rows: [{ Amount: 1 }] };
    const options = { name: "Amounts", recordId: { prefix: "  ", padding: 4 } };
    const preview = await codeOf(() => importedTableSchema(table, options));
    const imported = await codeOf(() => importTable(database, table, options));
    database.close();

    expect(preview).toBeDefined();
    expect(imported).toBe(preview);
  });

  it("never trips on the columns a reader produces", async () => {
    // Readers make their headers unique before anything reaches the import, so
    // a file that repeats a header arrives as distinct columns.
    const table = parseCsvTable("Amount,Amount,amount\n1,2,3\n");

    expect(new Set(table.columns).size).toBe(table.columns.length);
    const database = await Database.create();
    const imported = importTable(database, table, {
      name: "Amounts",
      recordId: { prefix: "A", padding: 4 },
    });
    database.close();
    expect(imported.columns).toHaveLength(3);
  });
});
