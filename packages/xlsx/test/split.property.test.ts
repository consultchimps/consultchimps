import { mkdtemp, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import {
  groupTableByColumn,
  type CellValue,
  type Table,
} from "@consultchimps/tabular";
import fc from "fast-check";
import { afterEach, describe, expect, it } from "vitest";
import * as XLSX from "xlsx";

import { splitWorkbookByColumn } from "../src/index.js";

const GROUP_COLUMN = "Group";
const ROW_ID_COLUMN = "RowId";
// A fixed seed keeps the suite deterministic, matching the repository's
// promise that identical inputs always produce identical results.
const runs = { numRuns: 300, seed: 20260805 };

const encoder = new TextEncoder();

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

const temporaryDirectories: string[] = [];

afterEach(async () => {
  await Promise.all(
    temporaryDirectories
      .splice(0)
      .map((directory) => rm(directory, { force: true, recursive: true })),
  );
});

describe("split output filenames", () => {
  it("caps a non-ASCII group value by UTF-8 bytes, not code points", async () => {
    const directory = await mkdtemp(
      path.join(tmpdir(), "consultchimps-xlsx-names-"),
    );
    temporaryDirectories.push(directory);
    const input = path.join(directory, "report.xlsx");
    const output = path.join(directory, "split");

    // Each of these code points encodes to three UTF-8 bytes, so a 100
    // character group value used to produce a 300 byte filename segment.
    const longValue = "中".repeat(100);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet([
        ["Region", "Amount"],
        [longValue, 1],
      ]),
      "Data",
    );
    await writeFile(
      input,
      XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }),
    );

    await splitWorkbookByColumn({
      column: "Region",
      input,
      outputDirectory: output,
    });

    const produced = await readdir(output);
    expect(produced).toHaveLength(1);

    const filename = produced[0]!;
    const segment = filename.slice("report-".length, -".xlsx".length);
    expect(encoder.encode(segment).length).toBeLessThanOrEqual(80);
    expect(encoder.encode(filename).length).toBeLessThanOrEqual(255);
    // 80 bytes divided by three bytes per character.
    expect([...segment]).toHaveLength(26);
  });
});
