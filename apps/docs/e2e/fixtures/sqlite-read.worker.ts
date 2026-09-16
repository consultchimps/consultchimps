import initialize from "@sqlite.org/sqlite-wasm";
import { openReadOnlySqlite } from "@consultchimps/db/sqlite-read";

const scope = self as unknown as {
  addEventListener(
    type: "message",
    listener: (event: MessageEvent<{ wasmUrl: string }>) => void,
  ): void;
  postMessage(value: unknown): void;
};

scope.addEventListener("message", (event) => {
  void run(event.data.wasmUrl).then(
    (result) => scope.postMessage({ ok: true, result }),
    (error: unknown) => scope.postMessage({ ok: false, error: String(error) }),
  );
});

async function run(wasmUrl: string): Promise<unknown> {
  const config = { locateFile: () => wasmUrl };
  const load = initialize as unknown as (
    options: typeof config,
  ) => ReturnType<typeof initialize>;
  const sqlite = await load(config);
  const source = new sqlite.oo1.DB(":memory:");
  source.exec(
    "CREATE TABLE example (id INTEGER, label TEXT); INSERT INTO example VALUES (9223372036854775807, 'North'), (2, 'South')",
  );
  const bytes = sqlite.capi.sqlite3_js_db_export(source.pointer!);
  source.close();
  const root = await navigator.storage.getDirectory();
  const names = async (): Promise<string[]> => {
    const result: string[] = [];
    for await (const [name] of root.entries()) result.push(name);
    return result.sort();
  };
  const before = await names();
  const calls: string[] = [];
  const reader = await openReadOnlySqlite(bytes, {
    runtime: {
      locateFile: (name) => {
        calls.push(name);
        return wasmUrl;
      },
    },
    maxRows: 1,
  });
  const refusal = (work: () => unknown): string => {
    try {
      work();
      return "unexpected-success";
    } catch (error) {
      return String((error as { code: unknown }).code);
    }
  };
  try {
    const wasmBinary = new Uint8Array(
      await (await fetch(wasmUrl)).arrayBuffer(),
    );
    const injected = await openReadOnlySqlite(bytes, {
      runtime: { wasmBinary },
    });
    injected.close();
    const id = String(
      reader.query("SELECT id FROM example WHERE label = ?", ["North"])
        .rows[0]![0],
    );
    const readonly = refusal(() => reader.query("DELETE FROM example"));
    const rowLimit = refusal(() => reader.query("SELECT * FROM example"));
    const limited = await openReadOnlySqlite(bytes, {
      runtime: { wasmBinary },
      maxSteps: 1000,
    });
    let steps: string;
    try {
      steps = refusal(() =>
        limited.query(
          "WITH RECURSIVE x(n) AS (VALUES(1) UNION ALL SELECT n+1 FROM x) SELECT sum(n) FROM x",
        ),
      );
    } finally {
      limited.close();
    }
    let badRuntime = "unexpected-success";
    try {
      await openReadOnlySqlite(bytes, {
        runtime: { wasmBinary: new Uint8Array([1]) },
      });
    } catch (error) {
      badRuntime = String((error as { code: unknown }).code);
    }
    const recovered = await openReadOnlySqlite(bytes, {
      runtime: { wasmBinary },
    });
    recovered.close();
    reader.close();
    reader.close();
    return {
      id,
      readonly,
      rowLimit,
      steps,
      badRuntime,
      calls,
      before,
      after: await names(),
      isolated: crossOriginIsolated,
      shared: typeof SharedArrayBuffer,
      memoryAfterClose: reader.wasmMemoryBytes,
    };
  } finally {
    reader.close();
  }
}
