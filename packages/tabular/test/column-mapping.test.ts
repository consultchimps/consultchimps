import { ConsultChimpsError } from "@consultchimps/core";
import { describe, expect, it } from "vitest";

import {
  applyColumnMapping,
  applyColumnMappingToTables,
  suggestColumnMapping,
  unionTables,
  validateColumnMapping,
  type ColumnMapping,
  type Table,
} from "../src/index.js";

/**
 * Capture the structured refusal a call is expected to throw, so a test can
 * assert on its stable code and machine-readable details.
 */
function refusalOf(run: () => unknown): ConsultChimpsError {
  try {
    run();
  } catch (error) {
    if (error instanceof ConsultChimpsError) {
      return error;
    }
    throw error;
  }
  throw new Error("Expected the call to be refused, but it returned.");
}

function invalidMappingProblem(document: unknown): string {
  const error = refusalOf(() => validateColumnMapping(document));
  expect(error.code).toBe("TABLE_MAPPING_INVALID");
  return String(error.details?.["problem"]);
}

const caseLog: Table = {
  columns: ["Case ID", "Region", "Opened On", "Total Amount"],
  rows: [
    {
      "Case ID": 1,
      Region: "north",
      "Opened On": "09/03/2024",
      "Total Amount": "1.234,50",
    },
  ],
  source: { file: "north.xlsx", firstDataRow: 2, sheet: "Cases" },
  sourceRows: [2],
};

describe("validateColumnMapping", () => {
  it("accepts a well-formed version 1 document", () => {
    const document = {
      version: 1,
      columns: [
        { name: "Case_ID", aliases: ["Case Number", "Reference"] },
        {
          name: "Opened_On",
          aliases: ["Opened On"],
          coercion: { type: "date", format: "DD/MM/YYYY" },
        },
        {
          name: "Total_Amount",
          aliases: ["Total"],
          coercion: {
            type: "number",
            decimalSeparator: ",",
            thousandsSeparator: ".",
          },
        },
      ],
      constants: { Dataset: "quarterly", Reviewed: false, Batch: 7 },
    };

    const mapping = validateColumnMapping(document);

    expect(mapping.version).toBe(1);
    expect(mapping.columns.map((column) => column.name)).toEqual([
      "Case_ID",
      "Opened_On",
      "Total_Amount",
    ]);
    expect(mapping.columns[0]?.aliases).toEqual(["Case Number", "Reference"]);
    expect(mapping.constants).toEqual({
      Batch: 7,
      Dataset: "quarterly",
      Reviewed: false,
    });
  });

  it("accepts a mapping with no aliases at all", () => {
    expect(
      validateColumnMapping({
        version: 1,
        columns: [{ name: "Case_ID", aliases: [] }],
      }).columns,
    ).toEqual([{ name: "Case_ID", aliases: [] }]);
  });

  it("names the first problem in a malformed document", () => {
    expect(invalidMappingProblem("not a mapping")).toBe("not_an_object");
    expect(invalidMappingProblem([])).toBe("not_an_object");
    expect(invalidMappingProblem({ version: 2, columns: [] })).toBe(
      "unsupported_version",
    );
    expect(invalidMappingProblem({ columns: [] })).toBe("unsupported_version");
    expect(invalidMappingProblem({ version: 1 })).toBe("columns_not_a_list");
    expect(invalidMappingProblem({ version: 1, columns: ["Case_ID"] })).toBe(
      "invalid_column_entry",
    );
    expect(
      invalidMappingProblem({ version: 1, columns: [{ aliases: [] }] }),
    ).toBe("missing_canonical_name");
    expect(
      invalidMappingProblem({ version: 1, columns: [{ name: "  " }] }),
    ).toBe("missing_canonical_name");
    expect(
      invalidMappingProblem({ version: 1, columns: [{ name: "Case_ID" }] }),
    ).toBe("aliases_not_a_list");
    expect(
      invalidMappingProblem({
        version: 1,
        columns: [{ name: "Case_ID", aliases: [""] }],
      }),
    ).toBe("invalid_alias");
  });

  it("reports the supported version alongside the rejected one", () => {
    const error = refusalOf(() =>
      validateColumnMapping({ version: 2, columns: [] }),
    );

    expect(error.message).toContain("version 1");
    expect(error.details).toMatchObject({ supportedVersion: 1, version: 2 });
  });

  it("rejects canonical columns that match the same source headers", () => {
    const error = refusalOf(() =>
      validateColumnMapping({
        version: 1,
        columns: [
          { name: "Case ID", aliases: [] },
          { name: "case_id", aliases: [] },
        ],
      }),
    );

    expect(error.details?.["problem"]).toBe("duplicate_canonical_column");
    expect(error.message).toContain("Case ID");
    expect(error.message).toContain("case_id");
  });

  it("rejects an alias claimed by two canonical columns", () => {
    const error = refusalOf(() =>
      validateColumnMapping({
        version: 1,
        columns: [
          { name: "Case_ID", aliases: ["Reference Number"] },
          { name: "Region", aliases: ["reference_number"] },
        ],
      }),
    );

    expect(error.details).toMatchObject({
      alias: "reference_number",
      canonicalColumn: "Region",
      conflictsWith: "Case_ID",
      problem: "duplicate_alias",
    });
  });

  it("rejects aliases of one canonical column that collide on normalized keys", () => {
    // "A+B" and "A-B" differ only in punctuation, so they normalize to one
    // key: the document cannot say which spelling it meant.
    const error = refusalOf(() =>
      validateColumnMapping({
        version: 1,
        columns: [{ name: "Checks", aliases: ["A+B", "A-B"] }],
      }),
    );

    expect(error.details).toMatchObject({
      alias: "A-B",
      canonicalColumn: "Checks",
      conflictsWith: "Checks",
      problem: "duplicate_alias",
    });
    expect(error.message).toContain("more than once");
  });

  it("accepts an alias that restates its own canonical column name", () => {
    // Spelling out the canonical column among its aliases is redundant but
    // unambiguous: both resolve to the same canonical column.
    expect(
      validateColumnMapping({
        version: 1,
        columns: [{ name: "Case_ID", aliases: ["Case ID", "Reference"] }],
      }).columns[0]?.aliases,
    ).toEqual(["Case ID", "Reference"]);
  });

  it("rejects unusable coercions", () => {
    const coerced = (coercion: unknown): unknown => ({
      version: 1,
      columns: [{ name: "Opened_On", aliases: [], coercion }],
    });

    expect(invalidMappingProblem(coerced("date"))).toBe("invalid_coercion");
    expect(invalidMappingProblem(coerced({ type: "currency" }))).toBe(
      "unknown_coercion_type",
    );
    expect(invalidMappingProblem(coerced({ type: "date" }))).toBe(
      "invalid_date_format",
    );
    // No year token, so the format cannot produce an ISO 8601 date.
    expect(
      invalidMappingProblem(coerced({ type: "date", format: "DD/MM" })),
    ).toBe("invalid_date_format");
    // Two day tokens make the parse ambiguous.
    expect(
      invalidMappingProblem(coerced({ type: "date", format: "DD/DD/YYYY" })),
    ).toBe("invalid_date_format");
    expect(
      invalidMappingProblem(
        coerced({
          type: "number",
          decimalSeparator: ",",
          thousandsSeparator: ",",
        }),
      ),
    ).toBe("invalid_number_separators");
    expect(
      invalidMappingProblem(coerced({ type: "number", decimalSeparator: "0" })),
    ).toBe("invalid_number_separators");
    expect(
      invalidMappingProblem(coerced({ type: "number", decimalSeparator: "" })),
    ).toBe("invalid_number_separators");
  });

  it("rejects unusable constant columns", () => {
    expect(
      invalidMappingProblem({ version: 1, columns: [], constants: [] }),
    ).toBe("constants_not_an_object");
    expect(
      invalidMappingProblem({
        version: 1,
        columns: [],
        constants: { Dataset: { nested: true } },
      }),
    ).toBe("invalid_constant_value");
    expect(
      invalidMappingProblem({
        version: 1,
        columns: [],
        constants: { "  ": "north" },
      }),
    ).toBe("invalid_constant_column");
  });

  it("rejects a constant column that collides with a mapped column", () => {
    const error = refusalOf(() =>
      validateColumnMapping({
        version: 1,
        columns: [{ name: "Region", aliases: ["Zone"] }],
        constants: { region: "north" },
      }),
    );

    expect(error.details).toMatchObject({
      conflictsWith: "Region",
      constantColumn: "region",
      problem: "constant_column_collision",
    });

    expect(
      invalidMappingProblem({
        version: 1,
        columns: [{ name: "Region", aliases: ["Zone"] }],
        constants: { zone: "north" },
      }),
    ).toBe("constant_column_collision");
  });
});

describe("applyColumnMapping", () => {
  it("folds alias spellings into the canonical column, written verbatim", () => {
    const mapping: ColumnMapping = {
      version: 1,
      columns: [
        { name: "Case_ID", aliases: [] },
        { name: "Region", aliases: ["Zone"] },
      ],
    };

    const spaced = applyColumnMapping(
      {
        columns: ["Case ID", "zone"],
        rows: [{ "Case ID": 1, zone: "north" }],
      },
      mapping,
    );
    const punctuated = applyColumnMapping(
      {
        columns: ["case-id", "Zone:"],
        rows: [{ "case-id": 2, "Zone:": "south" }],
      },
      mapping,
    );

    expect(spaced.table.columns).toEqual(["Case_ID", "Region"]);
    expect(spaced.table.rows).toEqual([{ Case_ID: 1, Region: "north" }]);
    expect(punctuated.table.columns).toEqual(["Case_ID", "Region"]);
    expect(punctuated.table.rows).toEqual([{ Case_ID: 2, Region: "south" }]);
    expect(spaced.unmappedColumns).toEqual([]);
  });

  it("passes unmapped columns through under their own names and reports them", () => {
    const result = applyColumnMapping(caseLog, {
      version: 1,
      columns: [{ name: "Case_ID", aliases: [] }],
    });

    expect(result.table.columns).toEqual([
      "Case_ID",
      "Region",
      "Opened On",
      "Total Amount",
    ]);
    expect(result.unmappedColumns).toEqual([
      "Region",
      "Opened On",
      "Total Amount",
    ]);
    expect(result.table.rows[0]).toMatchObject({
      Case_ID: 1,
      Region: "north",
    });
  });

  it("keeps the input table untouched", () => {
    const columns = [...caseLog.columns];
    const rows = caseLog.rows.map((row) => ({ ...row }));

    applyColumnMapping(caseLog, {
      version: 1,
      columns: [{ name: "Identifier", aliases: ["Case ID"] }],
    });

    expect(caseLog.columns).toEqual(columns);
    expect(caseLog.rows).toEqual(rows);
  });

  it("refuses when two columns of one table fold into one canonical column", () => {
    const error = refusalOf(() =>
      applyColumnMapping(
        {
          columns: ["Case ID", "Case_ID"],
          rows: [{ "Case ID": 1, Case_ID: 2 }],
          source: { file: "north.xlsx", sheet: "Cases" },
        },
        { version: 1, columns: [{ name: "Identifier", aliases: ["Case ID"] }] },
      ),
    );

    expect(error.code).toBe("TABLE_MAPPING_COLUMN_COLLISION");
    expect(error.message).toContain('sheet "Cases" of "north.xlsx"');
    expect(error.details).toMatchObject({
      canonicalColumn: "Identifier",
      columns: ["Case ID", "Case_ID"],
      file: "north.xlsx",
      sheet: "Cases",
    });
  });

  it("refuses when an alias and the canonical spelling appear in one table", () => {
    const error = refusalOf(() =>
      applyColumnMapping(
        {
          columns: ["Region", "Zone"],
          rows: [{ Region: "north", Zone: "south" }],
        },
        { version: 1, columns: [{ name: "Region", aliases: ["Zone"] }] },
      ),
    );

    expect(error.code).toBe("TABLE_MAPPING_COLUMN_COLLISION");
    expect(error.details).toMatchObject({
      canonicalColumn: "Region",
      columns: ["Region", "Zone"],
    });
  });

  it("treats reserved property names as ordinary columns", () => {
    // Written through JSON so the fixture holds real own properties: an
    // object literal with a "__proto__" key would set the prototype instead.
    const row = JSON.parse(
      '{"__proto__": 5, "constructor": "north", "Case ID": 1}',
    ) as Record<string, never>;

    const result = applyColumnMapping(
      {
        columns: ["__proto__", "constructor", "Case ID"],
        rows: [row],
      },
      {
        version: 1,
        columns: [{ name: "Identifier", aliases: ["Case ID"] }],
        constants: { toString: "quarterly" },
      },
    );

    expect(result.table.columns).toEqual([
      "__proto__",
      "constructor",
      "Identifier",
      "toString",
    ]);
    const mapped = result.table.rows[0];
    expect(Object.keys(mapped ?? {})).toEqual([
      "__proto__",
      "constructor",
      "Identifier",
      "toString",
    ]);
    expect(mapped?.["__proto__"]).toBe(5);
    expect(mapped?.["constructor"]).toBe("north");
    expect(mapped?.["toString"]).toBe("quarterly");
    // The row carries data, not an altered prototype chain.
    expect(Object.getPrototypeOf(mapped)).toBeNull();
    expect(Object.getPrototypeOf({})).toBe(Object.prototype);
  });

  it("reads reserved property names as missing when a row omits them", () => {
    const result = applyColumnMapping(
      { columns: ["__proto__"], rows: [{}] },
      { version: 1, columns: [{ name: "Marker", aliases: ["__proto__"] }] },
    );

    expect(result.table.rows[0]?.["Marker"]).toBeNull();
  });

  it("preserves the table's provenance", () => {
    const result = applyColumnMapping(caseLog, {
      version: 1,
      columns: [{ name: "Identifier", aliases: ["Case ID"] }],
    });

    expect(result.table.source).toEqual(caseLog.source);
    expect(result.table.sourceRows).toEqual(caseLog.sourceRows);
  });
});

describe("column mapping coercions", () => {
  const dateMapping: ColumnMapping = {
    version: 1,
    columns: [
      {
        name: "Opened_On",
        aliases: ["Opened On"],
        coercion: { type: "date", format: "DD/MM/YYYY" },
      },
    ],
  };

  it("parses declared date formats into ISO 8601 date strings", () => {
    const result = applyColumnMapping(
      {
        columns: ["Opened On"],
        rows: [
          { "Opened On": "09/03/2024" },
          { "Opened On": "29/02/2024" },
          { "Opened On": "" },
          { "Opened On": null },
        ],
      },
      dateMapping,
    );

    expect(result.table.rows).toEqual([
      { Opened_On: "2024-03-09" },
      { Opened_On: "2024-02-29" },
      { Opened_On: null },
      { Opened_On: null },
    ]);
  });

  it("accepts single-digit day and month tokens when they are declared", () => {
    const result = applyColumnMapping(
      { columns: ["Opened On"], rows: [{ "Opened On": "9/3/2024" }] },
      {
        version: 1,
        columns: [
          {
            name: "Opened_On",
            aliases: ["Opened On"],
            coercion: { type: "date", format: "D/M/YYYY" },
          },
        ],
      },
    );

    expect(result.table.rows).toEqual([{ Opened_On: "2024-03-09" }]);
  });

  it("refuses a value that is not a date in the declared format", () => {
    const error = refusalOf(() =>
      applyColumnMapping(
        {
          columns: ["Opened On"],
          rows: [{ "Opened On": "09/03/2024" }, { "Opened On": "2024-03-09" }],
          source: { file: "north.xlsx", firstDataRow: 2, sheet: "Cases" },
        },
        dateMapping,
      ),
    );

    expect(error.code).toBe("TABLE_MAPPING_COERCION_FAILED");
    expect(error.message).toContain("Row 3");
    expect(error.message).toContain('"DD/MM/YYYY"');
    expect(error.details).toMatchObject({
      canonicalColumn: "Opened_On",
      coercion: "date",
      column: "Opened On",
      file: "north.xlsx",
      row: 3,
      sheet: "Cases",
    });
  });

  it("refuses a date that does not exist in the calendar", () => {
    const error = refusalOf(() =>
      applyColumnMapping(
        { columns: ["Opened On"], rows: [{ "Opened On": "30/02/2023" }] },
        dateMapping,
      ),
    );

    expect(error.code).toBe("TABLE_MAPPING_COERCION_FAILED");
    expect(error.details).toMatchObject({ row: 2 });
  });

  it("parses numbers written with the declared separators", () => {
    const result = applyColumnMapping(
      {
        columns: ["Total"],
        rows: [
          { Total: "1.234,50" },
          { Total: "-12,5" },
          { Total: 42 },
          { Total: "" },
        ],
      },
      {
        version: 1,
        columns: [
          {
            name: "Total_Amount",
            aliases: ["Total"],
            coercion: {
              type: "number",
              decimalSeparator: ",",
              thousandsSeparator: ".",
            },
          },
        ],
      },
    );

    expect(result.table.rows).toEqual([
      { Total_Amount: 1234.5 },
      { Total_Amount: -12.5 },
      { Total_Amount: 42 },
      { Total_Amount: null },
    ]);
  });

  it("defaults to the English separators", () => {
    const result = applyColumnMapping(
      { columns: ["Total"], rows: [{ Total: "1,234.5" }] },
      {
        version: 1,
        columns: [
          {
            name: "Total_Amount",
            aliases: ["Total"],
            coercion: { type: "number" },
          },
        ],
      },
    );

    expect(result.table.rows).toEqual([{ Total_Amount: 1234.5 }]);
  });

  it("refuses a number whose separators sit where the format does not allow", () => {
    const grouped: ColumnMapping = {
      version: 1,
      columns: [
        {
          name: "Total_Amount",
          aliases: ["Total"],
          coercion: {
            type: "number",
            decimalSeparator: ",",
            thousandsSeparator: ".",
          },
        },
      ],
    };
    const parse = (written: string): unknown =>
      applyColumnMapping(
        { columns: ["Total"], rows: [{ Total: written }] },
        grouped,
      ).table.rows[0]?.["Total_Amount"];

    // Thousands separators may only group the integer part in threes. Before
    // the grouping was checked, "1,23.4" lost its separators and read as
    // 1.234 - a corrupted amount rather than a refusal.
    for (const malformed of ["1,23.4", "1.23,4", "1.2345,6", "12.34", "1..2"]) {
      expect(
        refusalOf(() => parse(malformed)).code,
        `expected "${malformed}" to be refused`,
      ).toBe("TABLE_MAPPING_COERCION_FAILED");
    }

    expect(parse("1.234.567,89")).toBe(1234567.89);
    expect(parse("123,45")).toBe(123.45);
    expect(parse("1234567,89")).toBe(1234567.89);
    expect(parse(",5")).toBe(0.5);
  });

  it("refuses a thousands separator inside the fractional part", () => {
    const error = refusalOf(() =>
      applyColumnMapping(
        { columns: ["Total"], rows: [{ Total: "1.5,000" }] },
        {
          version: 1,
          columns: [
            {
              name: "Total_Amount",
              aliases: ["Total"],
              coercion: { type: "number" },
            },
          ],
        },
      ),
    );

    expect(error.code).toBe("TABLE_MAPPING_COERCION_FAILED");
  });

  it("refuses a value that is not a number in the declared form", () => {
    const error = refusalOf(() =>
      applyColumnMapping(
        { columns: ["Total"], rows: [{ Total: "12 units" }] },
        {
          version: 1,
          columns: [
            {
              name: "Total_Amount",
              aliases: ["Total"],
              coercion: { type: "number" },
            },
          ],
        },
      ),
    );

    expect(error.code).toBe("TABLE_MAPPING_COERCION_FAILED");
    expect(error.details).toMatchObject({
      canonicalColumn: "Total_Amount",
      coercion: "number",
      column: "Total",
      row: 2,
    });
    // The refusal locates the cell without repeating its contents.
    expect(error.message).not.toContain("12 units");
  });
});

describe("column mapping constants", () => {
  const mapping: ColumnMapping = {
    version: 1,
    columns: [{ name: "Case_ID", aliases: ["Case ID"] }],
    constants: { Dataset: "quarterly", Reviewed: false },
  };

  it("appends constant columns after the mapped and unmapped columns", () => {
    const result = applyColumnMapping(
      {
        columns: ["Case ID", "Region"],
        rows: [{ "Case ID": 1, Region: "east" }],
      },
      mapping,
    );

    expect(result.table.columns).toEqual([
      "Case_ID",
      "Region",
      "Dataset",
      "Reviewed",
    ]);
    expect(result.table.rows).toEqual([
      { Case_ID: 1, Region: "east", Dataset: "quarterly", Reviewed: false },
    ]);
  });

  it("refuses a constant column that collides with a passed-through column", () => {
    const error = refusalOf(() =>
      applyColumnMapping(
        {
          columns: ["Case ID", "DATASET"],
          rows: [{ "Case ID": 1, DATASET: "east" }],
          source: { sheet: "Cases" },
        },
        mapping,
      ),
    );

    expect(error.code).toBe("TABLE_MAPPING_CONSTANT_COLLISION");
    expect(error.message).toContain('sheet "Cases"');
    expect(error.details).toMatchObject({
      column: "DATASET",
      constantColumn: "Dataset",
      sheet: "Cases",
    });
  });
});

describe("applyColumnMappingToTables", () => {
  const mapping: ColumnMapping = {
    version: 1,
    columns: [{ name: "Case_ID", aliases: ["Reference"] }],
  };

  const tables: Table[] = [
    {
      columns: ["Case ID", "Region"],
      rows: [{ "Case ID": 1, Region: "north" }],
      source: { file: "north.xlsx", sheet: "Cases" },
    },
    {
      columns: ["reference", "Region", "Owner"],
      rows: [{ reference: 2, Region: "south", Owner: "east team" }],
      source: { file: "south.xlsx", sheet: "Cases" },
    },
  ];

  it("maps every table and reports each distinct unmapped spelling once", () => {
    const result = applyColumnMappingToTables(tables, mapping);

    expect(result.tables.map((table) => table.columns)).toEqual([
      ["Case_ID", "Region"],
      ["Case_ID", "Region", "Owner"],
    ]);
    expect(result.unmappedColumns).toEqual(["Region", "Owner"]);
  });

  it("feeds a union that still records provenance", () => {
    const result = applyColumnMappingToTables(tables, mapping);
    const union = unionTables(result.tables);

    expect(union.columns).toEqual([
      "Case_ID",
      "Region",
      "Owner",
      "_source_file",
      "_source_sheet",
      "_source_row",
    ]);
    expect(union.rows).toEqual([
      {
        Case_ID: 1,
        Region: "north",
        Owner: null,
        _source_file: "north.xlsx",
        _source_sheet: "Cases",
        _source_row: 2,
      },
      {
        Case_ID: 2,
        Region: "south",
        Owner: "east team",
        _source_file: "south.xlsx",
        _source_sheet: "Cases",
        _source_row: 2,
      },
    ]);
  });

  it("leaves the union's source-column refusal in place for a canonical name that collides", () => {
    const mapped = applyColumnMappingToTables(
      [{ columns: ["Origin"], rows: [{ Origin: "north.xlsx" }] }],
      { version: 1, columns: [{ name: "_source_file", aliases: ["Origin"] }] },
    );

    expect(mapped.tables[0]?.columns).toEqual(["_source_file"]);
    expect(refusalOf(() => unionTables(mapped.tables)).code).toBe(
      "TABLE_SOURCE_COLUMN_COLLISION",
    );
    expect(
      unionTables(mapped.tables, { addSourceColumns: false }).columns,
    ).toEqual(["_source_file"]);
  });

  it("honours custom source column names", () => {
    const mapped = applyColumnMappingToTables(tables, mapping);
    const union = unionTables(mapped.tables, {
      sourceColumnNames: { file: "Origin", row: "Origin_Row", sheet: "Tab" },
    });

    expect(union.columns).toEqual([
      "Case_ID",
      "Region",
      "Owner",
      "Origin",
      "Origin_Row",
      "Tab",
    ]);
  });
});

describe("suggestColumnMapping", () => {
  const sources = [
    {
      columns: ["Case ID", "Region", "Failed Checks", "Timestamp"],
      file: "north.xlsx",
      sheet: "Cases",
    },
    {
      columns: ["Case_ID", "Region", "failed_checks", "Run Time"],
      file: "south.xlsx",
      sheet: "Cases",
    },
    {
      columns: ["case id", "Region"],
      file: "east.xlsx",
      sheet: "Cases",
    },
  ];

  it("groups spelling variants under the first spelling seen", () => {
    const suggestion = suggestColumnMapping(sources);

    expect(
      suggestion.groups.map((group) => [group.canonical, group.spellings]),
    ).toEqual([
      ["Case ID", ["Case ID", "Case_ID", "case id"]],
      ["Failed Checks", ["Failed Checks", "failed_checks"]],
    ]);
    expect(suggestion.groups[0]?.key).toBe("case_id");
  });

  it("drafts a version 1 mapping whose canonical names match every spelling", () => {
    const suggestion = suggestColumnMapping(sources);

    expect(suggestion.mapping).toEqual({
      version: 1,
      columns: [
        { name: "Case ID", aliases: [] },
        { name: "Failed Checks", aliases: [] },
      ],
    });
    // The draft is valid on its own and folds every grouped spelling.
    expect(validateColumnMapping(suggestion.mapping)).toEqual(
      suggestion.mapping,
    );
    expect(
      applyColumnMapping(
        { columns: ["case id", "failed_checks"], rows: [] },
        suggestion.mapping,
      ).table.columns,
    ).toEqual(["Case ID", "Failed Checks"]);
  });

  it("carries the evidence for every occurrence of a grouped spelling", () => {
    const suggestion = suggestColumnMapping(sources);

    expect(suggestion.groups[0]?.occurrences).toEqual([
      { column: "Case ID", file: "north.xlsx", sheet: "Cases" },
      { column: "Case_ID", file: "south.xlsx", sheet: "Cases" },
      { column: "case id", file: "east.xlsx", sheet: "Cases" },
    ]);
  });

  it("never groups synonyms that share no spelling", () => {
    const suggestion = suggestColumnMapping(sources);
    const grouped = suggestion.groups.flatMap((group) => group.spellings);

    // "Timestamp" and "Run Time" describe the same field but normalize to
    // different keys: joining them stays a manual mapping entry.
    expect(grouped).not.toContain("Timestamp");
    expect(grouped).not.toContain("Run Time");
  });

  it("leaves out columns that are spelled the same way everywhere", () => {
    expect(
      suggestColumnMapping(sources).groups.map((group) => group.canonical),
    ).not.toContain("Region");
  });

  it("groups variants inside a single input as well", () => {
    const suggestion = suggestColumnMapping([
      { columns: ["Failed Checks", "Failed_Checks"], sheet: "Cases" },
    ]);

    expect(suggestion.groups).toEqual([
      {
        canonical: "Failed Checks",
        key: "failed_checks",
        spellings: ["Failed Checks", "Failed_Checks"],
        occurrences: [
          { column: "Failed Checks", sheet: "Cases" },
          { column: "Failed_Checks", sheet: "Cases" },
        ],
      },
    ]);
  });

  it("is deterministic and follows the input order", () => {
    expect(suggestColumnMapping(sources)).toEqual(
      suggestColumnMapping(sources),
    );
    expect(
      suggestColumnMapping([...sources].reverse()).groups.map(
        (group) => group.canonical,
      ),
    ).toEqual(["case id", "failed_checks"]);
  });

  it("returns an empty draft when no input headers are given", () => {
    expect(suggestColumnMapping([])).toEqual({
      groups: [],
      mapping: { version: 1, columns: [] },
    });
  });
});
