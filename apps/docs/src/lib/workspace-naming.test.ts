import { describe, expect, it } from "vitest";

import {
  workspaceRecordIdPrefix,
  workspaceWorkingCopyName,
} from "./workspace-naming";

describe("workspace naming", () => {
  it("preserves ASCII record prefixes", () => {
    expect(workspaceRecordIdPrefix("Inventory Records")).toBe("INVENTOR");
  });

  it("takes eight whole code points before uppercasing a record prefix", () => {
    expect(workspaceRecordIdPrefix("ABCDEF𐐨GHI")).toBe("ABCDEF𐐀G");
    expect(workspaceRecordIdPrefix(`A${"𐐨".repeat(4)}tail`)).toBe(
      `A${"𐐀".repeat(4)}TAI`,
    );
    expect(workspaceRecordIdPrefix("AAAAAAAßtail")).toBe("AAAAAAASS");
  });

  it("keeps a supplementary character only when it fits the 80-unit stem", () => {
    const uniqueId = "00000000-0000-4000-8000-000000000001";
    const fitting = workspaceWorkingCopyName(
      `${"A".repeat(78)}𐐀.sqlite`,
      uniqueId,
    );
    const crossing = workspaceWorkingCopyName(
      `${"A".repeat(79)}𐐀.sqlite`,
      uniqueId,
    );

    expect(fitting).toBe(`${"A".repeat(78)}𐐀-${uniqueId}.sqlite`);
    expect(crossing).toBe(`${"A".repeat(79)}-${uniqueId}.sqlite`);
  });

  it("preserves DuckDB extension and the ASCII fallback", () => {
    expect(workspaceWorkingCopyName("...duckdb", "fixed")).toBe(
      "database-fixed.duckdb",
    );
  });
});
