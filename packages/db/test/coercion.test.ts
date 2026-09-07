import type { CellValue } from "@consultchimps/tabular";
import { describe, expect, it } from "vitest";

import {
  cellFromSqlValue,
  sqlValueFromCell,
  type ColumnType,
  type SqlValueType,
} from "../src/index.js";

// The single value-coercion contract, exercised as a matrix. Every write path
// (insertRecord, and the Table bridge through it) routes through
// sqlValueFromCell, so this table is the one place the type rules are pinned.

const accepted: Array<[ColumnType, CellValue, SqlValueType]> = [
  ["boolean", true, 1],
  ["boolean", false, 0],
  ["boolean", 1, 1],
  ["boolean", 0, 0],
  ["boolean", "true", 1],
  ["boolean", "Yes", 1],
  ["boolean", "0", 0],
  ["boolean", "no", 0],
  ["boolean", "", null],
  ["boolean", null, null],
  ["integer", 5, 5],
  ["integer", "5", 5],
  ["integer", 0, 0],
  ["integer", true, 1],
  ["integer", "", null],
  ["integer", null, null],
  ["real", 1.5, 1.5],
  ["real", "1.5", 1.5],
  ["real", "", null],
  ["real", null, null],
  ["text", "a", "a"],
  ["text", 5, "5"],
  ["text", null, null],
  ["date", "2020-01-01", "2020-01-01"],
  ["date", null, null],
];

const rejected: Array<[ColumnType, CellValue, string]> = [
  ["boolean", 2, "DB_INVALID_BOOLEAN"],
  ["boolean", -1, "DB_INVALID_BOOLEAN"],
  ["boolean", Number.NaN, "DB_INVALID_BOOLEAN"],
  ["boolean", Number.POSITIVE_INFINITY, "DB_INVALID_BOOLEAN"],
  ["boolean", "maybe", "DB_INVALID_BOOLEAN"],
  ["integer", 1.9, "DB_INVALID_NUMBER"],
  ["integer", "2.5", "DB_INVALID_NUMBER"],
  ["integer", "9007199254740993", "DB_INVALID_NUMBER"],
  ["integer", "x", "DB_INVALID_NUMBER"],
  ["integer", Number.NaN, "DB_INVALID_NUMBER"],
  ["integer", Number.POSITIVE_INFINITY, "DB_INVALID_NUMBER"],
  ["real", "x", "DB_INVALID_NUMBER"],
  ["real", Number.NaN, "DB_INVALID_NUMBER"],
  ["real", Number.POSITIVE_INFINITY, "DB_INVALID_NUMBER"],
];

describe("sqlValueFromCell", () => {
  it.each(accepted)("stores %s value %o as %o", (type, value, expected) => {
    expect(sqlValueFromCell(type, value)).toEqual(expected);
  });

  it.each(rejected)(
    "rejects %s value %o with %s and no leaked value",
    (type, value, code) => {
      try {
        sqlValueFromCell(type, value);
        expect.unreachable("value should have been rejected");
      } catch (error) {
        const themed = error as {
          code: string;
          message: string;
          details: Record<string, unknown>;
        };
        expect(themed.code).toBe(code);
        expect(themed.details).not.toHaveProperty("value");
        expect(themed.message).not.toContain(String(value));
      }
    },
  );
});

describe("cellFromSqlValue", () => {
  it("reads stored values back to their tabular type", () => {
    expect(cellFromSqlValue("boolean", 1)).toBe(true);
    expect(cellFromSqlValue("boolean", 0)).toBe(false);
    expect(cellFromSqlValue("boolean", null)).toBeNull();
    expect(cellFromSqlValue("integer", 7)).toBe(7);
    expect(cellFromSqlValue("real", 1.5)).toBe(1.5);
    expect(cellFromSqlValue("text", "a")).toBe("a");
    expect(cellFromSqlValue("date", "2020-01-01")).toBe("2020-01-01");
  });
});
