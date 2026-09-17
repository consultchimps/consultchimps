import path from "node:path";
import { beforeAll, describe, expect, it } from "vitest";
import { fetchFixture } from "../scripts/fetch-fixtures.js";
import { readPbiTables } from "../src/pipeline.js";
import {
  COMMITTED,
  FETCHED,
  canonical,
  escapeToken,
  fixturePath,
  oracleIndex,
  oracleRows,
  readFixture,
  tableDigest,
} from "./oracle.js";
import type { OracleFixture } from "./oracle.js";

/**
 * The corpus oracle. Every fixture decodes through `readPbiTables` and is
 * projected into the pbixray dump's canonical token grammar: cell for cell
 * against the committed dumps, and digest-level against the committed per-table
 * digests for any fetched file that is present. Fixtures whose file is absent
 * are reported by name rather than silently passing.
 */

const index = oracleIndex();
/** CI runs the whole corpus: a missing fixture is fetched, never skipped. */
const inCi = process.env.CI !== undefined && process.env.CI !== "";
// The refusal case carries no tables and has its own test.
const candidates: OracleFixture[] = index.fixtures.filter(
  (fixture) => fixture.tables.length > 0,
);
const present: OracleFixture[] = [];
const skipped: string[] = [];

beforeAll(async () => {
  for (const fixture of candidates) {
    if (fixturePath(fixture) !== null) {
      present.push(fixture);
      continue;
    }
    if (!inCi) {
      skipped.push(fixture.name);
      continue;
    }
    // A pinned URL and a committed digest, verified before anything is written,
    // so fetching here cannot introduce a fixture the repository did not pick.
    await fetchFixture(fixture, path.join(FETCHED, fixture.fileName));
    present.push(fixture);
  }
}, 600_000);

// A generous ceiling: the corpus is read whole, and the default peak bound is
// sized for an export, not for holding every table of every fixture at once.
const READ_OPTIONS = {
  includeHiddenTables: true,
  inputBytes: 64 * 1024 * 1024,
  decodedBytes: 256 * 1024 * 1024,
  peakBytes: 2 * 1024 * 1024 * 1024,
} as const;

describe("corpus oracle", () => {
  it("reports every fixture that is not available locally", () => {
    // Not a failure outside CI: a contributor without the cache runs the one
    // committed fixture. `node scripts/fetch-fixtures.ts` adds the rest.
    expect(skipped.every((name) => !COMMITTED.has(name))).toBe(true);
    if (skipped.length > 0)
      process.stdout.write(`corpus fixtures skipped: ${skipped.join(", ")}\n`);
    if (inCi) expect(skipped).toEqual([]);
    expect(present.length).toBeGreaterThanOrEqual(1);
  });

  for (const fixture of candidates) {
    it(`decodes ${fixture.name} exactly as the independent reader does`, async ({
      skip,
    }) => {
      const found = fixturePath(fixture);
      // Outside CI a fixture nobody fetched is skipped by name, never passed.
      if (found === null) skip(`${fixture.fileName} is not present locally`);
      const file = found!;
      const model = await readPbiTables(readFixture(file), READ_OPTIONS);
      const byName = new Map(model.tables.map((table) => [table.name, table]));
      expect([...byName.keys()].sort()).toEqual(
        fixture.tables.map((table) => table.name).sort(),
      );
      const cellLevel = COMMITTED.has(fixture.name);
      for (const expected of fixture.tables) {
        const table = byName.get(expected.name)!;
        expect(table.rowCount, `${expected.name} row count`).toBe(
          expected.rows,
        );
        expect(
          table.columns.map((column) => column.name),
          `${expected.name} columns`,
        ).toEqual([...expected.columns]);
        const rows: string[][] = [];
        for (let row = 0; row < table.rowCount; row++) {
          const cells: string[] = [];
          for (const column of table.columns)
            cells.push(
              escapeToken(canonical(column.values[row]!, column.type)),
            );
          rows.push(cells);
        }
        expect(
          tableDigest([...expected.columns], rows),
          `${expected.name} digest`,
        ).toBe(expected.digest);
        if (!cellLevel) continue;
        const dump = oracleRows(fixture.name, expected.name);
        expect(dump, `${expected.name} dump`).not.toBeNull();
        expect(dump!.columns).toEqual([...expected.columns]);
        expect(dump!.rows.length).toBe(rows.length);
        for (let row = 0; row < rows.length; row++)
          for (let column = 0; column < dump!.columns.length; column++)
            expect(
              rows[row]![column],
              `${fixture.name} ${expected.name} row ${row} column ${dump!.columns[column]}`,
            ).toBe(dump!.rows[row]![column]);
      }
    }, 300_000);
  }
});
