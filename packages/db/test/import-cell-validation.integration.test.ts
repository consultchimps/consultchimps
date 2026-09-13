import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, expect, test } from "vitest";

import { engineOf, inspectDatabase, valueAsBigInt } from "../src/database.js";
import { assertValidImportDateCell } from "../src/import/date-cell.js";
import {
  applyImport,
  prepareImport,
  resolveImport,
} from "../src/import/operations.js";
import type {
  ImportCell,
  ImportRecipe,
  ImportSource,
} from "../src/import/types.js";
import {
  createDatabase,
  createPreparedImport,
  exportDatabase,
  openDatabase,
  openPreparedImport,
} from "../src/node.js";
import {
  PREPARED_BINDING_TABLE,
  PREPARED_CAPTURE_TABLE,
  PREPARED_ROW_TABLE,
  preparedEngineOf,
} from "../src/prepared.js";

const directories: string[] = [];

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

const recipe: ImportRecipe = {
  version: 1,
  routes: [
    {
      source: "submission",
      selection: "Dates",
      destination: {
        kind: "new-table-infer",
        name: "Dates",
        recordId: { prefix: "DAT", padding: 6 },
      },
      columns: [],
    },
  ],
};

function source(cells: Readonly<Record<string, ImportCell>>): ImportSource {
  const bytes = new TextEncoder().encode(JSON.stringify(cells));
  return {
    key: "submission",
    readerVersion: "synthetic-cells-1",
    bytes: {
      name: "submission.xlsx",
      size: bytes.byteLength,
      async readAt(offset, length) {
        return bytes.slice(offset, offset + length);
      },
    },
    selections: [
      {
        key: "Dates",
        label: "Dates",
        async open() {
          return {
            columns: ["When", "Moment", "Amount"],
            async *batches() {
              yield [{ sourceRow: 2, cells }];
            },
            async close() {},
          };
        },
      },
    ],
  };
}

async function expectEmptyPlan(
  prepared: Parameters<typeof preparedEngineOf>[0],
): Promise<void> {
  const engine = preparedEngineOf(prepared);
  for (const table of [
    PREPARED_ROW_TABLE,
    PREPARED_CAPTURE_TABLE,
    PREPARED_BINDING_TABLE,
  ]) {
    const rows = await engine.query(`SELECT COUNT(*) AS count FROM ${table}`);
    expect(valueAsBigInt(rows[0]?.["count"], "row count")).toBe(0n);
  }
}

test("date cells accept the XLSX serial and worksheet date spellings", () => {
  const values: ImportCell[] = [
    { kind: "date", raw: "0", iso: "1904-01-01" },
    { kind: "date", raw: "6.1E1", iso: "1900-03-01" },
    {
      kind: "date",
      raw: "2024-01-02",
      iso: "2024-01-02T00:00:00.000Z",
    },
    {
      kind: "date",
      raw: "2024-01-02 03:04:05.1+04:00",
      iso: "2024-01-01T23:04:05.1Z",
    },
    {
      kind: "date",
      raw: "0000-01-01T00:00:00Z",
      iso: "0000-01-01T00:00:00Z",
    },
    {
      kind: "date",
      raw: "0.5",
      iso: "1904-01-01T12:00:00.000",
    },
    {
      kind: "formula",
      formula: "TODAY()",
      cached: { kind: "date", raw: "1", iso: "1904-01-02" },
    },
  ];
  for (const value of values) {
    expect(() => assertValidImportDateCell(value, "source")).not.toThrow();
  }

  for (const value of [
    { kind: "date", raw: "not-a-date", iso: "2024-01-02" },
    { kind: "date", raw: "1e999", iso: "2024-01-02" },
    { kind: "date", raw: "2024-02-30", iso: "2024-02-30" },
    { kind: "date", raw: "2024-01-02", iso: "2024-01-03" },
    {
      kind: "date",
      raw: "2024-01-02T03:04:05Z",
      iso: "2024-01-02T03:04:05+01:00",
    },
    {
      kind: "date",
      raw: "2024-01-02T03:04:05.1234Z",
      iso: "2024-01-02T03:04:05.123Z",
    },
  ] satisfies ImportCell[]) {
    expect(() => assertValidImportDateCell(value, "source")).toThrowError(
      expect.objectContaining({ code: "DB_INVALID_SOURCE_DATE" }),
    );
  }
});

for (const format of ["sqlite", "duckdb"] as const) {
  test(`${format}: a custom reader owns its numeric date epoch across saved-plan reopening`, async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "cc-reader-epoch-"));
    directories.push(directory);
    const { database } = await createDatabase({
      path: path.join(directory, `workspace.${format}`),
      format,
    });
    const customRecipe: ImportRecipe = {
      version: 1,
      routes: [
        {
          source: "submission",
          selection: "Dates",
          destination: {
            kind: "new-table",
            schema: {
              name: "Dates",
              columns: [
                { name: "When", type: "date" },
                { name: "Moment", type: "text" },
                { name: "Amount", type: "integer" },
              ],
              recordId: { prefix: "DAT", padding: 6 },
            },
          },
          columns: [
            { source: "When", target: "When", type: "date" },
            { source: "Moment", target: "Moment", type: "text" },
            { source: "Amount", target: "Amount", type: "integer" },
          ],
        },
      ],
    };
    const planPath = path.join(directory, "review.ccplan");
    const prepared = await createPreparedImport({
      path: planPath,
      database,
      recipe: customRecipe,
      baselineRevision: 0n,
    });
    const decodedDate = {
      kind: "date",
      raw: "946684800000",
      iso: "2000-01-01",
    } as const;
    const reader = source({
      When: decodedDate,
      Moment: { kind: "formula", cached: decodedDate },
      Amount: { kind: "number", raw: "1" },
    });
    try {
      await prepareImport({
        database,
        prepared,
        recipe: customRecipe,
        sources: [
          {
            ...reader,
            readerVersion: "synthetic-unix-milliseconds-1",
            bytes: { ...reader.bytes, name: "events.json" },
          },
        ],
      });
      const approved = await resolveImport({
        database,
        prepared,
        decisions: [],
      });
      if (approved.state !== "ready") throw new Error("Fixture needs review");
      await prepared.close();
      const reopened = await openPreparedImport({ path: planPath });
      try {
        const result = await applyImport({
          database,
          prepared: reopened,
          approved,
          requestId: "custom-reader-epoch",
        });
        expect(result.metrics.rowsImported).toBe(1);
        expect(
          await engineOf(database).query(
            'SELECT CAST("When" AS VARCHAR) AS decoded_date, "Moment" AS source_token FROM "Dates"',
          ),
        ).toEqual([
          { decoded_date: "2000-01-01", source_token: "946684800000" },
        ]);
      } finally {
        await reopened.close();
      }
    } finally {
      await prepared.close();
      await database.close();
    }
  });

  test(`${format}: malformed source and stored scalar cells never mutate the database`, async () => {
    const directory = await mkdtemp(path.join(tmpdir(), "cc-import-cells-"));
    directories.push(directory);
    const { database } = await createDatabase({
      path: path.join(directory, `workspace.${format}`),
      format,
    });
    const prepared = await createPreparedImport({
      path: path.join(directory, "review.ccplan"),
      database,
      recipe,
      baselineRevision: 0n,
    });
    const validCells = {
      When: {
        kind: "date",
        raw: "2024-01-02",
        iso: "2024-01-02T00:00:00Z",
      },
      Moment: {
        kind: "date",
        raw: "2024-01-02T03:04:05+00:00",
        iso: "2024-01-02T03:04:05+00:00",
      },
      Amount: { kind: "number", raw: "+.5E+1" },
    } satisfies Readonly<Record<string, ImportCell>>;
    try {
      await expect(
        prepareImport({
          database,
          prepared,
          recipe,
          sources: [
            source({
              ...validCells,
              When: {
                kind: "date",
                raw: "not-a-date",
                iso: "2024-01-02",
              },
            }),
          ],
        }),
      ).rejects.toMatchObject({
        code: "DB_INVALID_SOURCE_DATE",
        details: {
          source: "submission",
          selection: "Dates",
          sourceRow: 2,
          column: "When",
        },
      });
      await expectEmptyPlan(prepared);

      for (const raw of ["", "   ", "0x10"]) {
        await expect(
          prepareImport({
            database,
            prepared,
            recipe,
            sources: [
              source({ ...validCells, Amount: { kind: "number", raw } }),
            ],
          }),
        ).rejects.toMatchObject({ code: "DB_INVALID_SOURCE_NUMBER" });
        await expectEmptyPlan(prepared);
      }

      const preparedOutcome = await prepareImport({
        database,
        prepared,
        recipe,
        sources: [source(validCells)],
      });
      expect(preparedOutcome.prepared.state).toBe("needs-review");
      const approved = await resolveImport({
        database,
        prepared,
        decisions: [],
      });
      if (approved.state !== "ready") {
        throw new Error("Expected the import to be ready");
      }
      const before = await inspectDatabase({ database });
      const preparedEngine = preparedEngineOf(prepared);

      await preparedEngine.execute(
        `UPDATE ${PREPARED_ROW_TABLE} SET values_json = ?`,
        [
          JSON.stringify({
            ...validCells,
            When: {
              kind: "date",
              raw: "2024-02-30",
              iso: "2024-02-30",
            },
          }),
        ],
      );
      await expect(
        applyImport({
          database,
          prepared,
          approved,
          requestId: `${format}-invalid-date`,
        }),
      ).rejects.toMatchObject({ code: "DB_INVALID_PREPARED_IMPORT" });
      expect(await inspectDatabase({ database })).toEqual(before);

      await preparedEngine.execute(
        `UPDATE ${PREPARED_ROW_TABLE} SET values_json = ?`,
        [
          JSON.stringify({
            ...validCells,
            Amount: { kind: "number", raw: " " },
          }),
        ],
      );
      await expect(
        applyImport({
          database,
          prepared,
          approved,
          requestId: `${format}-invalid-number`,
        }),
      ).rejects.toMatchObject({ code: "DB_INVALID_PREPARED_IMPORT" });
      expect(await inspectDatabase({ database })).toEqual(before);

      await preparedEngine.execute(
        `UPDATE ${PREPARED_ROW_TABLE} SET values_json = ?`,
        [JSON.stringify(validCells)],
      );
      const result = await applyImport({
        database,
        prepared,
        approved,
        requestId: `${format}-valid-cells`,
      });
      expect(result.metrics.rowsImported).toBe(1);
      const rows = await engineOf(database).query(
        'SELECT CAST("When" AS VARCHAR) AS imported_date, CAST("Moment" AS VARCHAR) AS imported_timestamp, "Amount" AS amount FROM "Dates"',
      );
      expect(rows).toEqual([
        {
          imported_date: "2024-01-02",
          imported_timestamp:
            format === "duckdb"
              ? "2024-01-02 03:04:05"
              : "2024-01-02T03:04:05.000Z",
          amount: "5",
        },
      ]);

      const convertedPath = path.join(
        directory,
        `converted.${format === "sqlite" ? "duckdb" : "sqlite"}`,
      );
      await exportDatabase({
        database,
        output: convertedPath,
        format: format === "sqlite" ? "duckdb" : "sqlite",
      });
      const converted = await openDatabase({ path: convertedPath });
      try {
        expect(
          await engineOf(converted).query(
            'SELECT CAST("When" AS VARCHAR) AS imported_date, CAST("Moment" AS VARCHAR) AS imported_timestamp FROM "Dates"',
          ),
        ).toEqual([
          {
            imported_date: "2024-01-02",
            imported_timestamp: "2024-01-02 03:04:05",
          },
        ]);
      } finally {
        await converted.close();
      }
    } finally {
      await prepared.close();
      await database.close();
    }
  });
}
