import { isConsultChimpsError } from "@consultchimps/core";
import {
  suggestColumnMapping,
  validateColumnMapping,
  type ColumnEquivalenceGroup,
} from "@consultchimps/tabular";
import { describe, expect, it } from "vitest";

import {
  groupEvidence,
  mappingSummary,
  parseColumnMapping,
  reviewedColumnMapping,
  reviewedColumnMappingText,
  serializeColumnMapping,
} from "./column-mapping";

/**
 * The draft the consolidate page reviews, drafted the way the library drafts
 * it: two exports of one log whose header spelling drifted apart.
 */
function driftedGroups(): ColumnEquivalenceGroup[] {
  return suggestColumnMapping([
    { columns: ["Case_ID", "Failed Checks"], file: "north.xlsx", sheet: "Log" },
    { columns: ["Case ID", "Failed_Checks"], file: "south.xlsx", sheet: "vF" },
  ]).groups;
}

function groupFor(key: string): ColumnEquivalenceGroup {
  const group = driftedGroups().find((candidate) => candidate.key === key);
  expect(group, `no drafted group has the key "${key}"`).toBeDefined();
  return group!;
}

describe("parseColumnMapping", () => {
  it("reads a version 1 document", () => {
    const mapping = parseColumnMapping(
      '{"version":1,"columns":[{"name":"Case_ID","aliases":["Reference"]}]}',
    );
    expect(mapping.columns).toEqual([
      { name: "Case_ID", aliases: ["Reference"] },
    ]);
  });

  it("refuses a file that is not JSON at all under a stable code", () => {
    // No engine ever sees this one, so the refusal is raised here, under the
    // code the file surface uses for the same failure.
    let thrown: unknown;
    try {
      parseColumnMapping("version: 1\n");
    } catch (error) {
      thrown = error;
    }
    expect(isConsultChimpsError(thrown)).toBe(true);
    expect(isConsultChimpsError(thrown) && thrown.code).toBe(
      "XLSX_MAPPING_FILE_INVALID",
    );
    expect(isConsultChimpsError(thrown) && thrown.message).toMatch(
      /not valid JSON/u,
    );
  });

  it("hands a parsed but unusable document to the engine's refusal", () => {
    // The stable code is what the page renders through its ordinary failure
    // rendering, so the refusal must arrive as the engine's error, not a
    // second opinion formed here.
    let thrown: unknown;
    try {
      parseColumnMapping(
        '{"version":1,"columns":[{"name":"Case ID","aliases":[]},{"name":"case_id","aliases":[]}]}',
      );
    } catch (error) {
      thrown = error;
    }
    expect(isConsultChimpsError(thrown)).toBe(true);
    expect(isConsultChimpsError(thrown) && thrown.code).toBe(
      "TABLE_MAPPING_INVALID",
    );
  });
});

describe("mappingSummary", () => {
  it("counts canonical columns and aliases", () => {
    expect(
      mappingSummary({
        version: 1,
        columns: [
          { name: "Case_ID", aliases: ["Reference", "Case Number"] },
          { name: "Amount", aliases: [] },
        ],
      }),
    ).toBe("2 canonical columns, 2 aliases");
  });

  it("counts one of each in the singular, and names constant columns", () => {
    expect(
      mappingSummary({
        version: 1,
        columns: [{ name: "Case_ID", aliases: ["Reference"] }],
        constants: { Dataset: "quarterly" },
      }),
    ).toBe("1 canonical column, 1 alias, 1 constant column");
  });
});

describe("groupEvidence", () => {
  it("counts the worksheets and workbooks a group was seen in", () => {
    expect(groupEvidence(groupFor("failed_checks"))).toBe(
      "Seen in 2 worksheets across 2 workbooks",
    );
  });

  it("counts worksheets alone when the headers carry no provenance", () => {
    const [group] = suggestColumnMapping([
      { columns: ["Failed Checks"] },
      { columns: ["Failed_Checks"] },
    ]).groups;
    expect(group && groupEvidence(group)).toBe("Seen in 1 worksheet");
  });
});

describe("reviewedColumnMapping", () => {
  it("leaves an accepted proposal without aliases", () => {
    // Every spelling in the group already normalizes to the canonical name's
    // own key, so an alias would repeat that key and be refused.
    const mapping = reviewedColumnMapping([groupFor("failed_checks")], {});
    expect(mapping).toEqual({
      version: 1,
      columns: [{ name: "Failed Checks", aliases: [] }],
    });
    expect(() => validateColumnMapping(mapping)).not.toThrow();
  });

  it("keeps the group reachable when the canonical column is renamed", () => {
    const group = groupFor("failed_checks");
    const mapping = reviewedColumnMapping([group], {
      [group.key]: "Checks Failed",
    });
    // The new name normalizes to a different key, so one spelling travels as
    // an alias to put the group's own key back into the document.
    expect(mapping.columns).toEqual([
      { name: "Checks Failed", aliases: ["Failed Checks"] },
    ]);
    expect(() => validateColumnMapping(mapping)).not.toThrow();
  });

  it("adds no alias when a rename keeps the same normalized key", () => {
    const group = groupFor("failed_checks");
    expect(
      reviewedColumnMapping([group], { [group.key]: "failed_checks" }).columns,
    ).toEqual([{ name: "failed_checks", aliases: [] }]);
  });

  it("falls back to the proposal when the field is emptied", () => {
    const group = groupFor("case_id");
    expect(
      reviewedColumnMapping([group], { [group.key]: "   " }).columns,
    ).toEqual([{ name: "Case_ID", aliases: [] }]);
  });
});

describe("reviewedColumnMappingText", () => {
  it("writes an indented document that reads back as a version 1 mapping", () => {
    const groups = driftedGroups();
    const text = reviewedColumnMappingText(groups, {});
    expect(text.endsWith("\n")).toBe(true);
    expect(text).toContain('\n  "version": 1');
    expect(parseColumnMapping(text).columns).toHaveLength(groups.length);
  });

  it("refuses a review that renamed two groups into one column", () => {
    const groups = driftedGroups();
    expect(() =>
      reviewedColumnMappingText(
        groups,
        Object.fromEntries(groups.map((group) => [group.key, "Reference"])),
      ),
    ).toThrow(/not usable/u);
  });
});

describe("serializeColumnMapping", () => {
  it("writes the same text for the same mapping", () => {
    const mapping = { version: 1 as const, columns: [] };
    expect(serializeColumnMapping(mapping)).toBe(
      serializeColumnMapping(mapping),
    );
    expect(serializeColumnMapping(mapping)).toBe(
      '{\n  "version": 1,\n  "columns": []\n}\n',
    );
  });
});
