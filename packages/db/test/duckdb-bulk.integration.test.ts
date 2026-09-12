import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, expect, test } from "vitest";

import { NodeDuckDbEngine } from "../src/engines/duckdb/node.js";

const directories: string[] = [];

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

async function createEngine(): Promise<NodeDuckDbEngine> {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-duckdb-bulk-"));
  directories.push(directory);
  return NodeDuckDbEngine.create(path.join(directory, "test.duckdb"));
}

test("DuckDB bulk insertion preserves typed values across appender flushes", async () => {
  const engine = await createEngine();
  try {
    await engine.execute(
      "CREATE TABLE typed_values (exact HUGEINT, amount DECIMAL(38,9), observed_on DATE, enabled BOOLEAN, payload BLOB, note VARCHAR, optional INTEGER)",
    );
    const rows = Array.from(
      { length: 5_000 },
      (_, index) =>
        [
          index === 0 ? 123456789012345678901234567890n : BigInt(index),
          index === 0
            ? "12345678901234567890123456789.123456789"
            : "1.250000000",
          index === 0 ? "2024-02-29" : "2025-01-01",
          index % 2 === 0,
          new Uint8Array([index % 256, 0, 255]),
          `Synthetic ${index}`,
          index === 0 ? null : index,
        ] as const,
    );

    await engine.bulkInsert({
      table: "typed_values",
      columns: [
        "exact",
        "amount",
        "observed_on",
        "enabled",
        "payload",
        "note",
        "optional",
      ],
      rows,
    });

    const result = await engine.query(
      "SELECT exact::VARCHAR AS exact, amount::VARCHAR AS amount, observed_on::VARCHAR AS observed_on, enabled, hex(payload) AS payload, note, optional, count(*) OVER () AS row_count FROM typed_values ORDER BY typed_values.exact DESC LIMIT 1",
    );
    expect(result).toEqual([
      {
        exact: "123456789012345678901234567890",
        amount: "12345678901234567890123456789.123456789",
        observed_on: "2024-02-29",
        enabled: true,
        payload: "0000FF",
        note: "Synthetic 0",
        optional: null,
        row_count: 5000n,
      },
    ]);
  } finally {
    await engine.close();
  }
});

test("DuckDB bulk insertion supports reordered and partial column lists", async () => {
  const engine = await createEngine();
  try {
    await engine.execute(
      "CREATE TABLE partial_values (number INTEGER, label VARCHAR, defaulted INTEGER DEFAULT 7)",
    );
    await engine.bulkInsert({
      table: "partial_values",
      columns: ["label", "number"],
      rows: [["Synthetic", 42]],
    });
    expect(
      await engine.query("SELECT number, label, defaulted FROM partial_values"),
    ).toEqual([{ number: 42, label: "Synthetic", defaulted: 7 }]);
  } finally {
    await engine.close();
  }
});

test("DuckDB appender errors and cancellation roll back their transaction", async () => {
  const engine = await createEngine();
  try {
    await engine.execute("CREATE TABLE rollback_values (value INTEGER UNIQUE)");
    await expect(
      engine.transaction(async (transaction) => {
        await transaction.bulkInsert({
          table: "rollback_values",
          columns: ["value"],
          rows: [[1], [2]],
        });
        await transaction.bulkInsert({
          table: "rollback_values",
          columns: ["value"],
          rows: [[2]],
        });
      }),
    ).rejects.toThrow();
    expect(
      await engine.query("SELECT count(*) AS count FROM rollback_values"),
    ).toEqual([{ count: 0n }]);

    const controller = new AbortController();
    controller.abort("test cancellation");
    await expect(
      engine.transaction(async (transaction) => {
        await transaction.bulkInsert({
          table: "rollback_values",
          columns: ["value"],
          rows: [[3]],
          signal: controller.signal,
        });
      }),
    ).rejects.toMatchObject({ code: "OPERATION_ABORTED" });
    expect(
      await engine.query("SELECT count(*) AS count FROM rollback_values"),
    ).toEqual([{ count: 0n }]);
  } finally {
    await engine.close();
  }
});
