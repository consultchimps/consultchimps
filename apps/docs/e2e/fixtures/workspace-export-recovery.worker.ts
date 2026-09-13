import type { RandomAccessFile } from "@consultchimps/core";
import {
  applyImport,
  applySchema,
  inspectDatabase,
  planSchema,
  prepareImport,
  type ImportRecipe,
  type ImportSource,
} from "@consultchimps/db";
import {
  configureBrowserDatabaseRuntime,
  type BrowserDatabaseRuntime,
} from "@consultchimps/db/browser";
import type { DatabaseFormat } from "@consultchimps/db/schema";

type FailureMode = "cancel" | "write";

interface ExportCase {
  readonly source: DatabaseFormat;
  readonly target: DatabaseFormat;
  readonly failure?: FailureMode;
}

interface RunRequest {
  readonly id: number;
  readonly type: "run";
  readonly origin: string;
  readonly runId: string;
}

class ControlledFile implements RandomAccessFile {
  #bytes: Uint8Array;
  #mutationTriggered = false;
  closeCalls = 0;
  readonly readLengths: number[] = [];

  constructor(
    readonly name: string,
    bytes: Uint8Array,
    private readonly failure: FailureMode | undefined,
    private readonly controller: AbortController,
  ) {
    this.#bytes = bytes.slice();
  }

  get size(): number {
    return this.#bytes.byteLength;
  }

  async readAt(offset: number, length: number): Promise<Uint8Array> {
    this.readLengths.push(length);
    return this.#bytes.slice(offset, offset + length);
  }

  async writeAt(offset: number, bytes: Uint8Array): Promise<void> {
    if (!this.#mutationTriggered && this.failure === "write") {
      this.#mutationTriggered = true;
      this.#write(offset, bytes.subarray(0, Math.max(1, bytes.length / 2)));
      throw new Error("Injected destination write failure");
    }
    this.#write(offset, bytes);
    if (!this.#mutationTriggered && this.failure === "cancel") {
      this.#mutationTriggered = true;
      this.controller.abort();
    }
  }

  async truncate(size: number): Promise<void> {
    const resized = new Uint8Array(size);
    resized.set(this.#bytes.subarray(0, size));
    this.#bytes = resized;
  }

  async close(): Promise<void> {
    this.closeCalls += 1;
  }

  bytes(): Uint8Array {
    return this.#bytes.slice();
  }

  #write(offset: number, bytes: Uint8Array): void {
    const end = offset + bytes.byteLength;
    if (end > this.#bytes.byteLength) {
      const expanded = new Uint8Array(end);
      expanded.set(this.#bytes);
      this.#bytes = expanded;
    }
    this.#bytes.set(bytes, offset);
  }
}

function gate(): { readonly promise: Promise<void>; resolve(): void } {
  let release = (): void => undefined;
  const promise = new Promise<void>((resolve) => {
    release = resolve;
  });
  return { promise, resolve: release };
}

class PausedWriteFile extends ControlledFile {
  readonly #writeStarted = gate();
  readonly #writeReleased = gate();
  #paused = false;

  constructor(name: string, bytes: Uint8Array) {
    super(name, bytes, undefined, new AbortController());
  }

  override async writeAt(offset: number, bytes: Uint8Array): Promise<void> {
    if (!this.#paused) {
      this.#paused = true;
      this.#writeStarted.resolve();
      await this.#writeReleased.promise;
    }
    await super.writeAt(offset, bytes);
  }

  waitForWrite(): Promise<void> {
    return this.#writeStarted.promise;
  }

  releaseWrite(): void {
    this.#writeReleased.resolve();
  }
}

function originalBytes(): Uint8Array {
  return Uint8Array.from(
    { length: 1024 * 1024 + 17 },
    (_, index) => index % 251,
  );
}

function bytesEqual(left: Uint8Array, right: Uint8Array): boolean {
  if (left.byteLength !== right.byteLength) return false;
  return left.every((byte, index) => byte === right[index]);
}

function errorDetails(error: unknown): {
  readonly code?: string;
  readonly message: string;
} {
  const code =
    typeof error === "object" && error !== null
      ? Reflect.get(error, "code")
      : undefined;
  return {
    ...(typeof code === "string" ? { code } : {}),
    message: error instanceof Error ? error.message : String(error),
  };
}

async function createSource(
  runtime: BrowserDatabaseRuntime,
  format: DatabaseFormat,
  runId: string,
) {
  const created = await runtime.createDatabase({
    name: `${runId}-source.${format}`,
    format,
    schema: {
      version: 1,
      tables: [
        {
          name: "SyntheticRecords",
          recordId: { prefix: "SYN", padding: 4 },
          columns: [{ name: "value", type: "text" }],
        },
      ],
    },
  });
  const recipe: ImportRecipe = {
    version: 1,
    routes: [
      {
        source: "synthetic-source",
        selection: "SyntheticRecords",
        destination: { kind: "existing-table", table: "SyntheticRecords" },
        columns: [{ source: "value", target: "value", type: "text" }],
      },
    ],
  };
  const sourceBytes = new TextEncoder().encode("synthetic export row");
  const source: ImportSource = {
    key: "synthetic-source",
    readerVersion: "export-recovery-e2e-1",
    bytes: {
      name: "synthetic-source.txt",
      size: sourceBytes.byteLength,
      async readAt(offset, length) {
        return sourceBytes.slice(offset, offset + length);
      },
    },
    selections: [
      {
        key: "SyntheticRecords",
        label: "Synthetic records",
        async open() {
          return {
            columns: ["value"],
            async *batches() {
              yield [
                {
                  sourceRow: 2,
                  cells: {
                    value: { kind: "string" as const, value: "North" },
                  },
                },
              ];
            },
            async close() {},
          };
        },
      },
    ],
  };
  const baseline = await inspectDatabase({ database: created.database });
  const prepared = await runtime.createPreparedImport({
    name: `${runId}-${format}-source.ccplan`,
    database: created.database,
    recipe,
    baselineRevision: baseline.revision,
  });
  try {
    const outcome = await prepareImport({
      database: created.database,
      prepared,
      recipe,
      sources: [source],
    });
    if (outcome.prepared.state !== "ready") {
      throw new Error("Synthetic browser import did not become ready");
    }
    await applyImport({
      database: created.database,
      prepared,
      approved: outcome.prepared,
      requestId: `${runId}-${format}-source`,
    });
  } finally {
    await prepared.close();
  }
  return created;
}

async function runExportCase(
  runtime: BrowserDatabaseRuntime,
  sources: Readonly<
    Record<DatabaseFormat, Awaited<ReturnType<typeof createSource>>>
  >,
  runId: string,
  index: number,
  testCase: ExportCase,
) {
  const before = originalBytes();
  const controller = new AbortController();
  const destination = new ControlledFile(
    `${runId}-${String(index)}.${testCase.target}`,
    before,
    testCase.failure,
    controller,
  );
  let failure: ReturnType<typeof errorDetails> | undefined;
  let output:
    | {
        readonly size: number;
        readonly inspectedFormat: DatabaseFormat;
        readonly tables: readonly string[];
        readonly rowCount: string;
      }
    | undefined;
  try {
    await runtime.exportDatabase({
      database: sources[testCase.source].database,
      name: destination.name,
      destination,
      format: testCase.target,
      overwrite: true,
      signal: controller.signal,
    });
    const imported = await runtime.importDatabase({
      name: `${runId}-verification-${String(index)}.${testCase.target}`,
      source: destination,
    });
    try {
      const inspection = await inspectDatabase({ database: imported });
      output = {
        size: destination.size,
        inspectedFormat: inspection.format,
        tables: inspection.tables.map((table) => table.name),
        rowCount: String(inspection.tables[0]?.rowCount ?? 0n),
      };
    } finally {
      await imported.close();
    }
  } catch (error) {
    failure = errorDetails(error);
  }
  const sourceInspection = await inspectDatabase({
    database: sources[testCase.source].database,
  });
  return {
    source: testCase.source,
    target: testCase.target,
    failureMode: testCase.failure ?? null,
    ...(failure === undefined ? {} : { failure }),
    ...(output === undefined ? {} : { output }),
    destinationPreserved: bytesEqual(destination.bytes(), before),
    destinationSize: destination.size,
    originalSize: before.byteLength,
    closeCalls: destination.closeCalls,
    maximumRead: Math.max(0, ...destination.readLengths),
    sourceFormat: sourceInspection.format,
    sourceTables: sourceInspection.tables.map((table) => table.name),
    sourceRowCount: String(sourceInspection.tables[0]?.rowCount ?? 0n),
  };
}

async function runSqliteConcurrencyCase(
  runtime: BrowserDatabaseRuntime,
  source: Awaited<ReturnType<typeof createSource>>,
  runId: string,
) {
  const secondHandle = await runtime.openDatabase({
    name: `${runId}-source.sqlite`,
  });
  const plan = await planSchema({
    database: secondHandle,
    schema: {
      version: 1,
      tables: [
        {
          name: "SyntheticRecords",
          recordId: { prefix: "SYN", padding: 4 },
          columns: [
            { name: "value", type: "text" },
            { name: "reviewed", type: "boolean", nullable: true },
          ],
        },
      ],
    },
  });
  if (plan.state !== "ready") {
    await secondHandle.close();
    throw new Error("The concurrency schema plan did not become ready");
  }
  const destination = new PausedWriteFile(
    `${runId}-concurrent.sqlite`,
    originalBytes(),
  );
  const exporting = runtime.exportDatabase({
    database: source.database,
    name: destination.name,
    destination,
    format: "sqlite",
    overwrite: true,
  });
  await destination.waitForWrite();
  const mutating = applySchema({ database: secondHandle, plan });
  const mutationCompletedWhilePaused = await Promise.race([
    mutating.then(
      () => true,
      () => true,
    ),
    new Promise<false>((resolve) => {
      setTimeout(() => resolve(false), 250);
    }),
  ]);
  destination.releaseWrite();
  try {
    await Promise.all([exporting, mutating]);
    const imported = await runtime.importDatabase({
      name: `${runId}-concurrent-verification.sqlite`,
      source: destination,
    });
    try {
      const [snapshot, current] = await Promise.all([
        inspectDatabase({ database: imported }),
        inspectDatabase({ database: secondHandle }),
      ]);
      return {
        mutationCompletedWhilePaused,
        snapshotColumns:
          snapshot.tables[0]?.schema.columns.map((column) => column.name) ?? [],
        snapshotRowCount: String(snapshot.tables[0]?.rowCount ?? 0n),
        currentColumns:
          current.tables[0]?.schema.columns.map((column) => column.name) ?? [],
        currentRowCount: String(current.tables[0]?.rowCount ?? 0n),
        destinationCloseCalls: destination.closeCalls,
      };
    } finally {
      await imported.close();
    }
  } finally {
    destination.releaseWrite();
    await secondHandle.close();
  }
}

async function run(request: RunRequest) {
  const runtime = await configureBrowserDatabaseRuntime({
    sqlite: {
      wasmUrl: `${request.origin}/database-wasm/sqlite3.wasm`,
      directory: `/consultchimps-export-recovery-${request.runId}`,
      initialCapacity: 16,
    },
    duckdb: {
      wasmUrl: `${request.origin}/database-wasm/duckdb-eh.wasm`,
      workerUrl: `${request.origin}/database-wasm/duckdb-browser-eh.worker.js`,
    },
    opfsDirectory: `consultchimps-export-recovery-${request.runId}`,
  });
  const sqlite = await createSource(runtime, "sqlite", request.runId);
  const duckdb = await createSource(runtime, "duckdb", request.runId);
  const sources = { sqlite, duckdb };
  const cases: readonly ExportCase[] = [
    { source: "sqlite", target: "sqlite", failure: "write" },
    { source: "sqlite", target: "sqlite", failure: "cancel" },
    { source: "duckdb", target: "duckdb", failure: "write" },
    { source: "duckdb", target: "duckdb", failure: "cancel" },
    { source: "sqlite", target: "duckdb" },
    { source: "sqlite", target: "duckdb", failure: "write" },
    { source: "duckdb", target: "sqlite" },
    { source: "duckdb", target: "sqlite", failure: "cancel" },
  ];
  try {
    const results = [];
    for (const [index, testCase] of cases.entries()) {
      results.push(
        await runExportCase(runtime, sources, request.runId, index, testCase),
      );
    }
    return {
      recoveries: results,
      concurrency: await runSqliteConcurrencyCase(
        runtime,
        sqlite,
        request.runId,
      ),
    };
  } finally {
    await Promise.allSettled([
      sqlite.database.close(),
      duckdb.database.close(),
    ]);
  }
}

const scope = globalThis as unknown as {
  onmessage: ((event: MessageEvent<RunRequest>) => void) | null;
  postMessage(message: unknown): void;
};
scope.onmessage = (event) => {
  if (event.data.type !== "run") return;
  void run(event.data).then(
    (result) => scope.postMessage({ id: event.data.id, ok: true, result }),
    (error: unknown) =>
      scope.postMessage({
        id: event.data.id,
        ok: false,
        error: errorDetails(error),
      }),
  );
};
