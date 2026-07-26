import { describe, expect, it } from "vitest";

import { unionTables, uniqueHeaders, type Table } from "../src/index.js";

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
