import { describe, expect, it } from "vitest";

import { workspaceWorkingCopyName } from "./workspace-naming";

describe("workspace naming", () => {
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
