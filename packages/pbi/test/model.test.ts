import { describe, expect, it } from "vitest";
import type { AbfImage } from "../src/abf.js";
import type { CatalogColumn, CatalogTable } from "../src/catalog.js";
import { assembleTable, decodeColumn } from "../src/model.js";
import type { ColumnOutcome, UnverifiedTally } from "../src/model.js";
import {
  allocateWorksheets,
  planParts,
  tableExclusion,
  validateTypedValues,
} from "../src/pipeline.js";
import { idf, idfmeta, numericDictionary } from "./vertipaq-bytes.js";

/**
 * Section C, the per-column exclusion rows and the table-eligibility rows.
 * Every failure here is one column or one table, never a thrown pipeline error.
 */

/** A backup image assembled from named members, with no trailer. */
function image(members: Record<string, Uint8Array>): {
  bytes: Uint8Array;
  backup: AbfImage;
} {
  let total = 0;
  for (const member of Object.values(members)) total += member.length;
  const bytes = new Uint8Array(total);
  const entries: AbfImage["members"][number][] = [];
  let offset = 0;
  for (const [fileName, member] of Object.entries(members)) {
    bytes.set(member, offset);
    entries.push({
      path: fileName,
      fileName,
      storagePath: fileName,
      size: member.length,
      offset,
    });
    offset += member.length;
  }
  return { bytes, backup: { members: entries, errorCode: false } };
}

function column(overrides: Partial<CatalogColumn> = {}): CatalogColumn {
  return {
    id: 1,
    name: "Amount",
    storagePosition: 0,
    dataType: 6,
    hidden: false,
    calculated: false,
    dax: undefined,
    dictionary: null,
    hierarchyIndex: null,
    idfs: [],
    baseId: 0n,
    magnitude: 1,
    rowCount: 4,
    ...overrides,
  };
}

function table(overrides: Partial<CatalogTable> = {}): CatalogTable {
  return {
    id: 1,
    name: "Sales",
    hidden: false,
    calculated: false,
    dax: undefined,
    rowCount: 4,
    columns: [],
    ...overrides,
  };
}

const tally = (): UnverifiedTally => new Map();

function outcome(
  values: ColumnOutcome["values"],
  overrides: Partial<ColumnOutcome> = {},
): ColumnOutcome {
  return {
    column: column(),
    type: "int64",
    values,
    excluded: undefined,
    ...overrides,
  };
}

describe("column decoding failures are one column each", () => {
  it("decodes a well-formed value-encoded column", () => {
    const { bytes, backup } = image({
      "c.idfmeta": idfmeta([{ records: 4 }]),
      "c.idf": idf([{ runs: [[3, 4] as const] }]),
    });
    const result = decodeColumn(
      bytes,
      backup,
      column({ idfs: ["c.idf"], hierarchyIndex: "c.hidx" }),
      4,
      tally(),
    );
    expect(result.excluded).toBeUndefined();
    expect(result.values).toEqual([3n, 3n, 3n, 3n]);
  });

  it("excludes an AMO data type it does not support", () => {
    const { bytes, backup } = image({ "c.idfmeta": idfmeta([{ records: 4 }]) });
    const result = decodeColumn(
      bytes,
      backup,
      column({ dataType: 99, idfs: ["c.idf"] }),
      4,
      tally(),
    );
    expect(result.excluded).toBe("PBI_COLUMN_UNSUPPORTED_ENCODING");
    expect(result.values).toBeUndefined();
  });

  it("excludes a dictionary type it does not support", () => {
    const unsupported = new Uint8Array(28);
    new DataView(unsupported.buffer).setInt32(0, 7, true);
    const { bytes, backup } = image({ "c.dictionary": unsupported });
    const result = decodeColumn(
      bytes,
      backup,
      column({ dictionary: "c.dictionary", idfs: ["c.idf"] }),
      4,
      tally(),
    );
    expect(result.excluded).toBe("PBI_COLUMN_UNSUPPORTED_ENCODING");
  });

  it("excludes a structurally inconsistent .idf as unreadable", () => {
    const { bytes, backup } = image({
      "c.idfmeta": idfmeta([{ records: 4 }]),
      "c.idf": idf([{ runs: [[1, 1] as const], declaredEntries: 4_000_000 }]),
    });
    const result = decodeColumn(
      bytes,
      backup,
      column({ idfs: ["c.idf"], hierarchyIndex: "c.hidx" }),
      4,
      tally(),
    );
    expect(result.excluded).toBe("PBI_COLUMN_UNREADABLE");
  });

  it("excludes an .idfmeta that claims more records than the table has rows", () => {
    const { bytes, backup } = image({
      "c.idfmeta": idfmeta([{ records: 9_007_199_254_740_991 }]),
      "c.idf": idf([{ runs: [[1, 1] as const] }]),
    });
    const result = decodeColumn(
      bytes,
      backup,
      column({ idfs: ["c.idf"], hierarchyIndex: "c.hidx" }),
      4,
      tally(),
    );
    expect(result.excluded).toBe("PBI_COLUMN_ROW_ALIGNMENT_UNRECOVERABLE");
  });

  it("excludes a missing member as unreadable rather than throwing", () => {
    const { bytes, backup } = image({});
    const result = decodeColumn(
      bytes,
      backup,
      column({ idfs: ["absent.idf"], hierarchyIndex: "c.hidx" }),
      4,
      tally(),
    );
    expect(result.excluded).toBe("PBI_COLUMN_UNREADABLE");
  });

  it("pushes no more nulls than the table has rows for a column with no storage", () => {
    const { bytes, backup } = image({ "c.idfmeta": idfmeta([{ records: 4 }]) });
    const result = decodeColumn(
      bytes,
      backup,
      column({ idfs: ["c.idf"] }),
      4,
      tally(),
    );
    expect(result.values).toEqual([null, null, null, null]);
  });

  it("keeps a dictionary value past the safe-integer range exact", () => {
    const { bytes, backup } = image({
      "c.dictionary": numericDictionary(0, [9_007_199_254_740_993n]),
      "c.idfmeta": idfmeta([{ records: 2, minDataId: 3 }]),
      "c.idf": idf([{ runs: [[3, 2] as const] }]),
    });
    const result = decodeColumn(
      bytes,
      backup,
      column({ dictionary: "c.dictionary", idfs: ["c.idf"] }),
      2,
      tally(),
    );
    expect(result.values).toEqual([
      9_007_199_254_740_993n,
      9_007_199_254_740_993n,
    ]);
  });
});

describe("semantic validation after a successful byte decode", () => {
  it("excludes a date column holding a non-finite serial", () => {
    const checked = validateTypedValues(
      outcome([1, Number.NaN, 3], { type: "dateTimeSerial" }),
    );
    expect(checked.excluded).toBe("PBI_COLUMN_UNREADABLE");
    expect(checked.values).toBeUndefined();
  });

  it("keeps a date column of finite serials and nulls", () => {
    const checked = validateTypedValues(
      outcome([1, null, 3], { type: "dateTimeSerial" }),
    );
    expect(checked.excluded).toBeUndefined();
  });

  it("excludes a binary column whose base64 would pass the cell limit", () => {
    const checked = validateTypedValues(
      outcome([new Uint8Array(24_574)], { type: "binary" }),
    );
    expect(checked.excluded).toBe("PBI_BINARY_CELL_TOO_LONG");
    expect(checked.values).toBeUndefined();
  });

  it("keeps a binary column at the limit", () => {
    const checked = validateTypedValues(
      outcome([new Uint8Array(24_573)], { type: "binary" }),
    );
    expect(checked.excluded).toBeUndefined();
  });
});

describe("the table's row count is the catalog's", () => {
  it("excludes a column whose decoded length differs, even a majority", () => {
    const source = table({ rowCount: 4 });
    const short = () => outcome([1n, 2n, 3n], { column: column({ id: 2 }) });
    const { table: assembled, reasons } = assembleTable(source, [
      outcome([1n, 2n, 3n, 4n]),
      short(),
      { ...short(), column: column({ id: 3 }) },
    ]);
    // Two of three columns agree on three rows. The catalog says four, so the
    // two are excluded and the one that matches survives.
    expect(assembled.rowCount).toBe(4);
    expect(assembled.columns.map((entry) => entry.id)).toEqual([1]);
    expect(reasons.get(2)).toBe("PBI_COLUMN_ROW_ALIGNMENT_UNRECOVERABLE");
    expect(reasons.get(3)).toBe("PBI_COLUMN_ROW_ALIGNMENT_UNRECOVERABLE");
  });

  it("leaves a table with no surviving column empty", () => {
    const { table: assembled } = assembleTable(table({ rowCount: 9 }), [
      outcome([1n]),
    ]);
    expect(assembled.columns).toHaveLength(0);
    expect(assembled.rowCount).toBe(9);
  });

  it("keeps a column exclusion the decoder already made", () => {
    const { reasons } = assembleTable(table({ rowCount: 1 }), [
      outcome(undefined, { excluded: "PBI_COLUMN_UNSUPPORTED_ENCODING" }),
    ]);
    expect(reasons.get(1)).toBe("PBI_COLUMN_UNSUPPORTED_ENCODING");
  });
});

describe("table eligibility", () => {
  it("skips a hidden table by default and includes it on request", () => {
    const hidden = table({ hidden: true });
    expect(tableExclusion(hidden, false)).toBe("PBI_TABLE_HIDDEN");
    expect(tableExclusion(hidden, true)).toBeUndefined();
  });

  it("refuses a table wider than a worksheet", () => {
    const wide = table({
      columns: Array.from({ length: 16_385 }, (_, index) =>
        column({ id: index }),
      ),
    });
    expect(tableExclusion(wide, false)).toBe("PBI_TABLE_TOO_WIDE");
    const atLimit = table({
      columns: Array.from({ length: 16_384 }, (_, index) =>
        column({ id: index }),
      ),
    });
    expect(tableExclusion(atLimit, false)).toBeUndefined();
  });

  it("applies the hidden policy before the column limit", () => {
    const both = table({
      hidden: true,
      columns: Array.from({ length: 16_385 }, (_, index) =>
        column({ id: index }),
      ),
    });
    expect(tableExclusion(both, false)).toBe("PBI_TABLE_HIDDEN");
  });
});

describe("worksheet allocation", () => {
  const synthetic = (rowCount: number, name = "Sales") => ({
    id: 1,
    name,
    hidden: false,
    calculated: false,
    rowCount,
    columns: [
      {
        name: "Amount",
        id: 1,
        storagePosition: 0,
        type: "int64" as const,
        hidden: false,
        values: [],
      },
    ],
  });

  it("splits one row past the worksheet limit and names the parts", () => {
    const { plans, parts, splitTables } = allocateWorksheets([
      synthetic(1_048_576),
    ]);
    expect(splitTables).toBe(1);
    expect(plans.map((plan) => plan.sheetName)).toEqual(["Sales", "Sales_2"]);
    expect(parts.get(1)).toEqual([
      { sheetName: "Sales", start: 0, end: 1_048_575 },
      { sheetName: "Sales_2", start: 1_048_575, end: 1_048_576 },
    ]);
  });

  it("keeps a table at the limit whole", () => {
    const { plans, splitTables } = allocateWorksheets([synthetic(1_048_575)]);
    expect(splitTables).toBe(0);
    expect(plans).toHaveLength(1);
  });

  it("gives a zero-row table one header-only part", () => {
    const { plans, parts } = allocateWorksheets([synthetic(0)]);
    expect(plans).toHaveLength(1);
    expect(plans[0]!.start).toBe(0);
    expect(plans[0]!.end).toBe(0);
    expect(parts.get(1)).toEqual([{ sheetName: "Sales", start: 0, end: 0 }]);
    expect(planParts(0)).toEqual([{ start: 0, end: 0 }]);
  });

  it("lets a later table take the next free suffix after a split", () => {
    const first = synthetic(1_048_576);
    const second = { ...synthetic(1, "Sales_2"), id: 2 };
    const { plans } = allocateWorksheets([first, second]);
    expect(plans.map((plan) => plan.sheetName)).toEqual([
      "Sales",
      "Sales_2",
      "Sales_2_2",
    ]);
  });
});
