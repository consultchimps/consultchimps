import { describe, expect, it } from "vitest";

import {
  columnKey,
  groupTableByColumn,
  normalizedColumnKey,
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

  it("does not generate a name the header row already carries", () => {
    // Counting occurrences left to right renames the second A to A_2 and then
    // meets the real A_2, and two columns end up with one name. As object keys
    // that means one column's values silently overwrite another's.
    const headers = uniqueHeaders(["A", "A", "A_2"]);
    expect(headers).toEqual(["A", "A_3", "A_2"]);
    expect(new Set(headers).size).toBe(3);
  });

  it("keeps a filled-in name clear of a real one that matches it", () => {
    // The first column to claim a name keeps it, so the blank at index 0 takes
    // "column_1" and the real header of that name is the one that moves.
    expect(uniqueHeaders([null, "column_1"])).toEqual([
      "column_1",
      "column_1_2",
    ]);
    expect(uniqueHeaders(["column_1", null])).toEqual(["column_1", "column_2"]);
  });

  it("names a long run of repeated headers in one pass", () => {
    // Every duplicate takes the next free suffix for its base rather than
    // re-probing from 2, so a wide row of one repeated header stays linear;
    // the names are the ones the slower search would have produced.
    const count = 20000;
    const result = uniqueHeaders(Array.from({ length: count }, () => "A"));
    expect(result[0]).toBe("A");
    expect(result[1]).toBe("A_2");
    expect(result[count - 1]).toBe(`A_${count}`);
    expect(new Set(result.map(columnKey)).size).toBe(count);
  });

  it("treats names that differ only by case as one name", () => {
    // The same rule every other column lookup here uses, so a header row cannot
    // name two columns that a later lookup would not tell apart.
    expect(uniqueHeaders(["Region", "region", "REGION"])).toEqual([
      "Region",
      "region_2",
      "REGION_3",
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

describe("normalizedColumnKey", () => {
  // Spelling variants of the kind different source systems produce when they
  // each export the same schema slightly differently.
  it("maps spacing, punctuation, and case variants to one key", () => {
    expect(normalizedColumnKey("Failed Checks")).toBe("failed_checks");
    expect(normalizedColumnKey("Failed_Checks")).toBe("failed_checks");
    expect(normalizedColumnKey("Failed  Checks ")).toBe("failed_checks");
    expect(normalizedColumnKey("Reviewer: Lead Contact")).toBe(
      "reviewer_lead_contact",
    );
    expect(normalizedColumnKey("Reviewer:Lead Contact")).toBe(
      "reviewer_lead_contact",
    );
    expect(normalizedColumnKey("Reviewer_Lead_Contact")).toBe(
      "reviewer_lead_contact",
    );
    expect(normalizedColumnKey("Modified_ON")).toBe("modified_on");
    expect(normalizedColumnKey("Modified On")).toBe("modified_on");
    expect(normalizedColumnKey("S.No.")).toBe("s_no");
  });

  it("keeps distinct names distinct and survives symbol-only headers", () => {
    expect(normalizedColumnKey("Client Name")).not.toBe(
      normalizedColumnKey("Name"),
    );
    expect(normalizedColumnKey("Total Checks")).not.toBe(
      normalizedColumnKey("Failed Checks"),
    );
    expect(normalizedColumnKey("###")).toBe("###");
  });

  it("folds case the same way whatever locale the host runs in", () => {
    // The locale-aware fold reads the host's default locale, where a Turkish
    // or Azeri locale folds "ID" to "ıd". Consolidation would then match
    // headers on one machine and not on another, so both keys use the
    // locale-independent fold and always produce "id".
    expect(columnKey("ID")).toBe("id");
    expect(normalizedColumnKey("Case ID")).toBe("case_id");
    expect(normalizedColumnKey("CASE_ID")).toBe("case_id");
    expect(columnKey("ID")).not.toBe("ID".toLocaleLowerCase("tr"));
    expect(normalizedColumnKey("İstanbul")).toBe(
      "İstanbul".toLowerCase().replace(/[^\p{L}\p{N}]+/gu, "_"),
    );
  });

  it("stays linear on long runs of separators", () => {
    const hostile = `${"_".repeat(50_000)}x${"_".repeat(50_000)}`;
    const start = performance.now();
    expect(normalizedColumnKey(hostile)).toBe("x");
    expect(performance.now() - start).toBeLessThan(1_000);
  });
});

describe("unionTables with normalizeHeaders", () => {
  const spaced: Table = {
    columns: ["Case_ID", "Failed Checks", "Reviewer: Lead Contact"],
    rows: [{ Case_ID: 1, "Failed Checks": 5, "Reviewer: Lead Contact": "A" }],
    source: { file: "spaced.xlsx", sheet: "Log" },
  };
  const underscored: Table = {
    columns: ["Case_ID", "Failed_Checks", "Reviewer_Lead_Contact"],
    rows: [{ Case_ID: 2, Failed_Checks: 7, Reviewer_Lead_Contact: "B" }],
    source: { file: "underscored.xlsx", sheet: "Log" },
  };

  it("merges header variants into the first-seen spelling", () => {
    const result = unionTables([spaced, underscored], {
      addSourceColumns: false,
      normalizeHeaders: true,
    });

    expect(result.columns).toEqual([
      "Case_ID",
      "Failed Checks",
      "Reviewer: Lead Contact",
    ]);
    expect(result.rows).toEqual([
      { Case_ID: 1, "Failed Checks": 5, "Reviewer: Lead Contact": "A" },
      { Case_ID: 2, "Failed Checks": 7, "Reviewer: Lead Contact": "B" },
    ]);
  });

  it("keeps the exact-match behaviour when the option is off", () => {
    const result = unionTables([spaced, underscored], {
      addSourceColumns: false,
    });

    expect(result.columns).toEqual([
      "Case_ID",
      "Failed Checks",
      "Reviewer: Lead Contact",
      "Failed_Checks",
      "Reviewer_Lead_Contact",
    ]);
  });

  it("resolves same-table collisions to the first column", () => {
    const result = unionTables(
      [
        {
          columns: ["Failed Checks", "Failed_Checks"],
          rows: [{ "Failed Checks": 5, Failed_Checks: 9 }],
        },
      ],
      { addSourceColumns: false, normalizeHeaders: true },
    );

    expect(result.columns).toEqual(["Failed Checks"]);
    expect(result.rows).toEqual([{ "Failed Checks": 5 }]);
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
