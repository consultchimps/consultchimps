import { describe, expect, it } from "vitest";
import type { PbiReasonCode } from "../src/errors.js";
import {
  buildWarnings,
  reasonEntries,
  serializeManifest,
} from "../src/manifest.js";
import type { PbiManifest, PbiManifestColumn } from "../src/manifest.js";
import { exclusionCounts } from "../src/pipeline.js";

/** ADR Decision 9: the manifest's bounds, order and warning vocabulary. */

function column(
  reasons: readonly (readonly [PbiReasonCode, number])[],
  overrides: Partial<PbiManifestColumn> = {},
): PbiManifestColumn {
  return {
    name: "Amount",
    id: 1,
    type: "int64",
    reasons: reasons.map(([code, count]) => ({ code, count })),
    ...overrides,
  };
}

const empty: PbiManifest = {
  schemaVersion: 1,
  tables: [],
  excludedTables: [],
  unverifiedPaths: [],
};

describe("reason entries", () => {
  it("sorts by code in ascending ASCII order and omits zero counts", () => {
    const entries = reasonEntries(
      new Map<PbiReasonCode, number>([
        ["PBI_TEXT_TRUNCATED", 2],
        ["PBI_DATE_ROUNDED", 0],
        ["PBI_BINARY_AS_BASE64", 7],
        ["PBI_NUMERIC_AS_TEXT", 1],
      ]),
    );
    expect(entries).toEqual([
      { code: "PBI_BINARY_AS_BASE64", count: 7 },
      { code: "PBI_NUMERIC_AS_TEXT", count: 1 },
      { code: "PBI_TEXT_TRUNCATED", count: 2 },
    ]);
  });

  it("returns nothing for a column with no conversions", () => {
    expect(reasonEntries(undefined)).toEqual([]);
    expect(reasonEntries(new Map())).toEqual([]);
  });
});

describe("warnings", () => {
  it("counts a column-scope code in columns and a values-scope code in values", () => {
    const manifest: PbiManifest = {
      ...empty,
      tables: [
        {
          name: "Sales",
          id: 1,
          reasons: [],
          parts: [],
          columns: [
            column([["PBI_NUMERIC_AS_TEXT", 1_000_000]]),
            column([["PBI_COLUMN_UNREADABLE", 1]], { id: 2 }),
          ],
        },
      ],
    };
    const warnings = buildWarnings(manifest, 0);
    // A million converted cells are one column's count, not a million entries.
    expect(warnings.some((line) => line.includes("1000000"))).toBe(true);
    expect(warnings.some((line) => line.startsWith("1 column"))).toBe(true);
  });

  it("emits one line per distinct code, in ascending ASCII code order", () => {
    const manifest: PbiManifest = {
      ...empty,
      tables: [
        {
          name: "Sales",
          id: 1,
          reasons: [],
          parts: [],
          columns: [
            column([
              ["PBI_TEXT_TRUNCATED", 3],
              ["PBI_DATE_AS_TEXT", 2],
              ["PBI_BINARY_AS_BASE64", 1],
            ]),
          ],
        },
      ],
      unverifiedPaths: [{ code: "PBI_UNVERIFIED_BOOLEAN_TYPE", count: 1 }],
    };
    const warnings = buildWarnings(manifest, 0);
    expect(warnings).toHaveLength(4);
    expect(warnings[0]).toContain("binary value");
    expect(warnings[1]).toContain("date value");
    expect(warnings[2]).toContain("text value");
    // The unverified codes sort after every reason code.
    expect(warnings[3]).toContain("true or false");
  });

  it("names the include-hidden option in the hidden-table warning", () => {
    const manifest: PbiManifest = {
      ...empty,
      excludedTables: [
        {
          name: "LocalDateTable_x",
          id: 2,
          reasons: [{ code: "PBI_TABLE_HIDDEN", count: 1 }],
          columns: [],
        },
      ],
    };
    // The cure is named in words every surface can use: an option name belongs
    // to a caller's interface, and this line is read in a terminal and in a
    // browser alike.
    expect(buildWarnings(manifest, 0)[0]).toContain("Export hidden tables");
    expect(buildWarnings(manifest, 0)[0]).not.toContain("includeHiddenTables");
  });

  it("counts split tables from the allocation, not from a column", () => {
    expect(buildWarnings(empty, 2)[0]).toContain("2 tables");
    expect(buildWarnings(empty, 2)[0]).toContain("split");
    expect(buildWarnings(empty, 0)).toEqual([]);
  });

  it("says nothing for DAX or worksheet provenance alone", () => {
    const manifest: PbiManifest = {
      ...empty,
      tables: [
        {
          name: "Calc",
          id: 1,
          dax: "EVALUATE 1",
          reasons: [],
          parts: [{ sheetName: "Calc", start: 0, end: 3 }],
          columns: [column([])],
        },
      ],
    };
    expect(buildWarnings(manifest, 0)).toEqual([]);
  });
});

describe("serialization", () => {
  it("writes a fixed property order with no timestamp and no source value", () => {
    const manifest: PbiManifest = {
      schemaVersion: 1,
      tables: [
        {
          name: "Sales",
          id: 1,
          dax: "EVALUATE 1",
          reasons: [{ code: "PBI_TABLE_SPLIT", count: 1 }],
          parts: [{ sheetName: "Sales", start: 0, end: 3 }],
          columns: [
            column([["PBI_BINARY_AS_BASE64", 2]], {
              encoding: "base64",
              dateEpoch: "1899-12-30",
            }),
          ],
        },
      ],
      excludedTables: [
        {
          name: "Hidden",
          id: 2,
          reasons: [{ code: "PBI_TABLE_HIDDEN", count: 1 }],
          columns: [],
        },
      ],
      unverifiedPaths: [{ code: "PBI_UNVERIFIED_MULTIPLE_SEGMENTS", count: 4 }],
    };
    const text = new TextDecoder().decode(serializeManifest(manifest));
    const parsed = JSON.parse(text) as Record<string, unknown>;
    expect(Object.keys(parsed)).toEqual([
      "schemaVersion",
      "tables",
      "excludedTables",
      "unverifiedPaths",
    ]);
    const table = (parsed.tables as Record<string, unknown>[])[0]!;
    expect(Object.keys(table)).toEqual([
      "name",
      "id",
      "dax",
      "reasons",
      "parts",
      "columns",
    ]);
    const entry = (table.columns as Record<string, unknown>[])[0]!;
    expect(Object.keys(entry)).toEqual([
      "name",
      "id",
      "type",
      "encoding",
      "dateEpoch",
      "reasons",
    ]);
    expect(text).not.toMatch(/\d{4}-\d{2}-\d{2}T/);
  });

  it("omits an absent dax, encoding and epoch rather than writing null", () => {
    const text = new TextDecoder().decode(
      serializeManifest({
        ...empty,
        tables: [
          {
            name: "Sales",
            id: 1,
            reasons: [],
            parts: [],
            columns: [column([])],
          },
        ],
      }),
    );
    expect(text).not.toContain("dax");
    expect(text).not.toContain("encoding");
    expect(text).not.toContain("null");
  });

  it("writes the same bytes for the same manifest", () => {
    expect(serializeManifest(empty)).toEqual(serializeManifest(empty));
  });
});

describe("the anonymous counts of a no-exportable-tables refusal", () => {
  it("counts an excluded column once, whatever its value count", () => {
    // One binary column with five oversized cells. The manifest keeps the five;
    // the refusal's aggregate is tables and columns, so it sees one column.
    const counts = exclusionCounts([
      {
        source: {
          id: 1,
          name: "Attachments",
          hidden: false,
          calculated: false,
          dax: undefined,
          rowCount: 5,
          columns: [],
        },
        reasons: ["PBI_TABLE_NO_EXPORTABLE_COLUMNS"],
        columns: [column([["PBI_BINARY_CELL_TOO_LONG", 5]])],
      },
    ]);
    expect(counts).toEqual([
      { code: "PBI_BINARY_CELL_TOO_LONG", tables: 0, columns: 1 },
      { code: "PBI_TABLE_NO_EXPORTABLE_COLUMNS", tables: 1, columns: 0 },
    ]);
  });

  it("adds a column per excluded column and a table per excluded table", () => {
    const table = (id: number, columns: PbiManifestColumn[]) => ({
      source: {
        id,
        name: `t${id}`,
        hidden: true,
        calculated: false,
        dax: undefined,
        rowCount: 0,
        columns: [],
      },
      reasons: ["PBI_TABLE_HIDDEN" as PbiReasonCode],
      columns,
    });
    const counts = exclusionCounts([
      table(1, [
        column([["PBI_COLUMN_UNREADABLE", 1]]),
        column([["PBI_COLUMN_UNREADABLE", 1]], { id: 2 }),
      ]),
      table(2, []),
    ]);
    expect(counts).toEqual([
      { code: "PBI_COLUMN_UNREADABLE", tables: 0, columns: 2 },
      { code: "PBI_TABLE_HIDDEN", tables: 2, columns: 0 },
    ]);
  });
});
