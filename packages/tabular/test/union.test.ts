import { describe, expect, it } from "vitest";

import {
  groupTableByColumn,
  unionTables,
  uniqueHeaders,
  type Table,
} from "../src/index.js";

describe("uniqueHeaders", () => {
  it("fills blank headers and disambiguates duplicates", () => {
    expect(uniqueHeaders(["Name", " name ", null, "Amount"])).toEqual([
      "Name",
      "name_2",
      "column_3",
      "Amount",
    ]);
  });
});

describe("unionTables", () => {
  it("unions case-insensitive columns and records provenance", () => {
    const tables: Table[] = [
      {
        columns: ["Client", "Amount"],
        rows: [{ Client: "A", Amount: 10 }],
        source: { file: "one.xlsx", firstDataRow: 4, sheet: "North" },
        sourceRows: [4],
      },
      {
        columns: ["client", "Status"],
        rows: [{ client: "B", Status: "Open" }],
        source: { file: "two.xlsx", firstDataRow: 2, sheet: "South" },
        sourceRows: [7],
      },
    ];

    const result = unionTables(tables);

    expect(result.columns).toEqual([
      "Client",
      "Amount",
      "Status",
      "_source_file",
      "_source_sheet",
      "_source_row",
    ]);
    expect(result.rows).toEqual([
      {
        Client: "A",
        Amount: 10,
        Status: null,
        _source_file: "one.xlsx",
        _source_sheet: "North",
        _source_row: 4,
      },
      {
        Client: "B",
        Amount: null,
        Status: "Open",
        _source_file: "two.xlsx",
        _source_sheet: "South",
        _source_row: 7,
      },
    ]);
  });
});

describe("groupTableByColumn", () => {
  it("groups typed values in first-seen order and preserves provenance", () => {
    const table: Table = {
      columns: ["Client", "Region"],
      rows: [
        { Client: "A", Region: "North" },
        { Client: "B", Region: 1 },
        { Client: "C", Region: "North" },
        { Client: "D", Region: "1" },
        { Client: "E", Region: " " },
      ],
      source: { file: "clients.xlsx", firstDataRow: 2, sheet: "Clients" },
      sourceRows: [2, 3, 4, 5, 6],
    };

    const result = groupTableByColumn(table, " region ");

    expect(result.column).toBe("Region");
    expect(result.skippedRows).toBe(0);
    expect(result.groups.map((group) => group.value)).toEqual([
      "North",
      1,
      "1",
      null,
    ]);
    expect(result.groups[0]?.table.rows).toEqual([
      { Client: "A", Region: "North" },
      { Client: "C", Region: "North" },
    ]);
    expect(result.groups[0]?.table.sourceRows).toEqual([2, 4]);
  });

  it("can skip blank values without silently losing the row count", () => {
    const result = groupTableByColumn(
      {
        columns: ["Client", "Region"],
        rows: [
          { Client: "A", Region: null },
          { Client: "B", Region: "" },
          { Client: "C", Region: "South" },
        ],
      },
      "Region",
      { includeBlank: false },
    );

    expect(result.groups.map((group) => group.value)).toEqual(["South"]);
    expect(result.skippedRows).toBe(2);
  });

  it("reports the available columns when the requested column is missing", () => {
    expect(() =>
      groupTableByColumn(
        {
          columns: ["Client", "Region"],
          rows: [{ Client: "A", Region: "North" }],
        },
        "Office",
      ),
    ).toThrowError(/Column "Office" was not found/);
  });
});
