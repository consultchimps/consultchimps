import fc from "fast-check";
import { describe, expect, it } from "vitest";

import {
  groupTableByColumn,
  type CellValue,
  type Table,
} from "../src/index.js";

const GROUP_COLUMN = "Group";
const ROW_ID_COLUMN = "RowId";
// A fixed seed keeps the suite deterministic, matching the repository's
// promise that identical inputs always produce identical results.
const runs = { numRuns: 300, seed: 20260805 };

// Values that exercise every branch of the grouping key: blanks, whitespace
// only strings, numbers, booleans and text that collides across types.
const groupValue: fc.Arbitrary<CellValue> = fc.oneof(
  fc.constant(null),
  fc.constantFrom("", " ", "  "),
  fc.constantFrom("north", "NORTH", "1", "true"),
  fc.integer({ min: -3, max: 3 }),
  fc.boolean(),
  fc.string({ maxLength: 8 }),
);

function tableOf(values: CellValue[]): Table {
  return {
    columns: [GROUP_COLUMN, ROW_ID_COLUMN],
    rows: values.map((value, index) => ({
      [GROUP_COLUMN]: value,
      [ROW_ID_COLUMN]: index,
    })),
  };
}

function rowIdsOf(table: Table): number[] {
  return table.rows.map((row) => row[ROW_ID_COLUMN] as number);
}

function isBlank(value: CellValue): boolean {
  return value === null || (typeof value === "string" && value.trim() === "");
}

describe("groupTableByColumn partitions rows", () => {
  it("places every row in exactly one group or counts it skipped", () => {
    fc.assert(
      fc.property(
        fc.array(groupValue, { maxLength: 40 }),
        fc.boolean(),
        (values, includeBlank) => {
          const table = tableOf(values);
          const grouped = groupTableByColumn(table, GROUP_COLUMN, {
            includeBlank,
          });

          const groupedIds = grouped.groups.flatMap((group) =>
            rowIdsOf(group.table),
          );
          const expectedIds = values
            .map((value, index) => ({ index, value }))
            .filter(({ value }) => includeBlank || !isBlank(value))
            .map(({ index }) => index);

          // No row is duplicated across groups or inside one.
          expect(new Set(groupedIds).size).toBe(groupedIds.length);
          // Grouped rows plus skipped rows account for the whole input.
          expect(groupedIds.length + grouped.skippedRows).toBe(
            table.rows.length,
          );
          // Exactly the rows that were not skipped are grouped.
          expect([...groupedIds].sort((a, b) => a - b)).toEqual(expectedIds);
        },
      ),
      runs,
    );
  });

  it("skips rows only when blank values are excluded", () => {
    fc.assert(
      fc.property(
        fc.array(groupValue, { maxLength: 40 }),
        fc.boolean(),
        (values, includeBlank) => {
          const grouped = groupTableByColumn(tableOf(values), GROUP_COLUMN, {
            includeBlank,
          });

          expect(grouped.skippedRows).toBe(
            includeBlank ? 0 : values.filter(isBlank).length,
          );
        },
      ),
      runs,
    );
  });

  it("gives every row in a group the same normalized value", () => {
    fc.assert(
      fc.property(fc.array(groupValue, { maxLength: 40 }), (values) => {
        const table = tableOf(values);
        const grouped = groupTableByColumn(table, GROUP_COLUMN);

        for (const group of grouped.groups) {
          for (const rowId of rowIdsOf(group.table)) {
            const original = values[rowId]!;
            expect(isBlank(original) ? null : original).toEqual(group.value);
          }
        }
      }),
      runs,
    );
  });

  it("is deterministic for identical inputs", () => {
    fc.assert(
      fc.property(fc.array(groupValue, { maxLength: 40 }), (values) => {
        expect(groupTableByColumn(tableOf(values), GROUP_COLUMN)).toEqual(
          groupTableByColumn(tableOf(values), GROUP_COLUMN),
        );
      }),
      runs,
    );
  });
});
