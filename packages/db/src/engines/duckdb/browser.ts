import {
  AsyncDuckDB,
  DuckDBAccessMode,
  DuckDBDataProtocol,
  VoidLogger,
  type AsyncDuckDBConnection,
} from "@duckdb/duckdb-wasm/dist/duckdb-browser";

import { throwIfAborted } from "@consultchimps/core";
import type { RandomAccessFile } from "@consultchimps/core";

import { copyBrowserFileToDestination } from "../../browser-file-copy.js";
import {
  exportReadonlyDuckDbSnapshot,
  type DuckDbSnapshotStorage,
} from "../../browser-duckdb-snapshot.js";
import type {
  DatabaseEngine,
  EngineRow,
  EngineTransaction,
  EngineValue,
} from "../../internal/engine.js";
import { quoteIdentifier } from "../../schema.js";

function normalizeValue(value: unknown): EngineValue {
  if (
    value === null ||
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "bigint" ||
    typeof value === "boolean" ||
    value instanceof Uint8Array
  ) {
    return value;
  }
  if (value instanceof Date) return value.toISOString();
  return String(value);
}

function normalizeRow(value: unknown): EngineRow {
  const source =
    typeof value === "object" && value !== null && "toJSON" in value
      ? (value as { toJSON(): unknown }).toJSON()
      : value;
  if (typeof source !== "object" || source === null || Array.isArray(source)) {
    throw new Error("DuckDB returned a row in an unsupported shape");
  }
  const row: Record<string, EngineValue> = Object.create(null);
  for (const [name, cell] of Object.entries(source)) {
    row[name] = normalizeValue(cell);
  }
  return row;
}

export interface BrowserDuckDbEngineOptions {
  readonly wasmUrl: string;
  readonly workerUrl: string;
  readonly storageName: string;
  readonly fileHandle: BrowserFileHandle;
  readonly walHandle: BrowserFileHandle;
  readonly readonly?: boolean | undefined;
  readonly createReadonlySnapshot?:
    (() => Promise<BrowserDuckDbSnapshotStorage>) | undefined;
}

export interface BrowserDuckDbSnapshotStorage extends DuckDbSnapshotStorage {
  readonly fileHandle: BrowserFileHandle;
  readonly walHandle: BrowserFileHandle;
}

interface BrowserFile {
  readonly size: number;
  slice(start?: number, end?: number): Blob;
}

export interface BrowserFileHandle {
  getFile(): Promise<BrowserFile>;
}

export class BrowserDuckDbEngine implements DatabaseEngine {
  readonly format = "duckdb" as const;
  readonly interruptible = true;
  readonly #database: AsyncDuckDB;
  readonly #connection: AsyncDuckDBConnection;
  readonly #fileHandle: BrowserFileHandle;
  readonly #readonly: boolean;
  readonly #createReadonlySnapshot:
    (() => Promise<BrowserDuckDbSnapshotStorage>) | undefined;
  #tail: Promise<void> = Promise.resolve();
  #closed = false;
  #closeRequested = false;
  #closeAttempt: Promise<void> | undefined;
  #checkpointAttempted = false;
  #connectionClosed = false;
  #filesFlushed = false;

  private constructor(
    database: AsyncDuckDB,
    connection: AsyncDuckDBConnection,
    fileHandle: BrowserFileHandle,
    readonly: boolean,
    createReadonlySnapshot:
      (() => Promise<BrowserDuckDbSnapshotStorage>) | undefined,
  ) {
    this.#database = database;
    this.#connection = connection;
    this.#fileHandle = fileHandle;
    this.#readonly = readonly;
    this.#createReadonlySnapshot = createReadonlySnapshot;
  }

  static async open(
    options: BrowserDuckDbEngineOptions,
  ): Promise<BrowserDuckDbEngine> {
    type DuckWorker = NonNullable<ConstructorParameters<typeof AsyncDuckDB>[1]>;
    const WorkerConstructor = (
      globalThis as unknown as {
        readonly Worker: new (
          url: string,
          options: { readonly type: "classic" },
        ) => DuckWorker;
      }
    ).Worker;
    const worker = new WorkerConstructor(options.workerUrl, {
      type: "classic",
    });
    const database = new AsyncDuckDB(new VoidLogger(), worker);
    try {
      await database.instantiate(options.wasmUrl);
      await database.registerFileHandle(
        options.storageName,
        options.fileHandle,
        DuckDBDataProtocol.BROWSER_FSACCESS,
        true,
      );
      await database.registerFileHandle(
        `${options.storageName}.wal`,
        options.walHandle,
        DuckDBDataProtocol.BROWSER_FSACCESS,
        true,
      );
      await database.open({
        path: options.storageName,
        accessMode:
          options.readonly === true
            ? DuckDBAccessMode.READ_ONLY
            : DuckDBAccessMode.READ_WRITE,
        useDirectIO: true,
        arrowLosslessConversion: true,
      });
      return new BrowserDuckDbEngine(
        database,
        await database.connect(),
        options.fileHandle,
        options.readonly === true,
        options.createReadonlySnapshot,
      );
    } catch (error) {
      await database.terminate().catch(() => undefined);
      throw error;
    }
  }

  async #exclusive<T>(work: () => Promise<T>, allowClose = false): Promise<T> {
    const previous = this.#tail;
    let release: () => void = () => undefined;
    this.#tail = new Promise<void>((resolve) => {
      release = resolve;
    });
    await previous;
    try {
      if (this.#closed || (this.#closeRequested && !allowClose)) {
        throw new Error("DuckDB engine is closed");
      }
      return await work();
    } finally {
      release();
    }
  }

  async #run(
    sql: string,
    values: readonly EngineValue[] = [],
  ): Promise<unknown> {
    if (values.length === 0) return this.#connection.query(sql);
    const statement = await this.#connection.prepare(sql);
    try {
      return await statement.query(
        ...values.map((value) =>
          typeof value === "bigint" ? value.toString() : value,
        ),
      );
    } finally {
      await statement.close();
    }
  }

  async #execute(
    sql: string,
    values: readonly EngineValue[] = [],
  ): Promise<void> {
    await this.#run(sql, values);
  }

  async #query(
    sql: string,
    values: readonly EngineValue[] = [],
  ): Promise<readonly EngineRow[]> {
    const table = await this.#run(sql, values);
    if (
      typeof table !== "object" ||
      table === null ||
      !("toArray" in table) ||
      typeof table.toArray !== "function"
    ) {
      throw new Error("DuckDB did not return a table");
    }
    return (table.toArray() as readonly unknown[]).map(normalizeRow);
  }

  async execute(sql: string, values?: readonly EngineValue[]): Promise<void> {
    await this.#exclusive(() => this.#execute(sql, values));
  }

  async query(
    sql: string,
    values?: readonly EngineValue[],
  ): Promise<readonly EngineRow[]> {
    return this.#exclusive(() => this.#query(sql, values));
  }

  async #bulkInsert(options: {
    readonly table: string;
    readonly columns: readonly string[];
    readonly rows: readonly (readonly EngineValue[])[];
    readonly signal?: AbortSignal | undefined;
  }): Promise<void> {
    if (options.rows.length === 0) return;
    const quotedColumns = options.columns.map(quoteIdentifier).join(", ");
    const rowSql = `(${options.columns.map(() => "?").join(", ")})`;
    const maximumRows = Math.max(1, Math.floor(1000 / options.columns.length));
    for (let offset = 0; offset < options.rows.length; offset += maximumRows) {
      throwIfAborted(options.signal, "db.bulk-insert");
      const batch = options.rows.slice(offset, offset + maximumRows);
      await this.#execute(
        `INSERT INTO ${quoteIdentifier(options.table)} (${quotedColumns}) VALUES ${batch
          .map(() => rowSql)
          .join(", ")}`,
        batch.flat(),
      );
    }
  }

  async bulkInsert(options: {
    readonly table: string;
    readonly columns: readonly string[];
    readonly rows: readonly (readonly EngineValue[])[];
    readonly signal?: AbortSignal | undefined;
  }): Promise<void> {
    await this.#exclusive(() => this.#bulkInsert(options));
  }

  async transaction<T>(
    work: (transaction: EngineTransaction) => Promise<T>,
  ): Promise<T> {
    return this.#exclusive(async () => {
      await this.#execute("BEGIN TRANSACTION");
      const transaction: EngineTransaction = {
        execute: (sql, values) => this.#execute(sql, values),
        query: (sql, values) => this.#query(sql, values),
        bulkInsert: (options) => this.#bulkInsert(options),
      };
      try {
        const result = await work(transaction);
        await this.#execute("COMMIT");
        return result;
      } catch (error) {
        await this.#execute("ROLLBACK");
        throw error;
      }
    });
  }

  async readTransaction<T>(
    work: (transaction: EngineTransaction) => Promise<T>,
  ): Promise<T> {
    return this.transaction(work);
  }

  async checkpoint(): Promise<void> {
    await this.#exclusive(async () => {
      if (!this.#readonly) await this.#execute("CHECKPOINT");
    });
  }

  interrupt(): void {
    void this.#connection.cancelSent();
  }

  async copyTo(
    destination: RandomAccessFile,
    signal?: AbortSignal,
  ): Promise<number> {
    return this.#exclusive(async () => {
      if (this.#readonly) {
        if (this.#createReadonlySnapshot === undefined) {
          throw new Error(
            "A read-only DuckDB browser engine requires snapshot storage to export.",
          );
        }
        await this.#database.flushFiles();
        throwIfAborted(signal, "db.export");
        return exportReadonlyDuckDbSnapshot({
          destination,
          signal,
          allocate: this.#createReadonlySnapshot,
          prepare: async (storage) => {
            const rows = await this.#query(
              "SELECT current_database() AS database_name",
            );
            const sourceName = rows[0]?.["database_name"];
            if (typeof sourceName !== "string") {
              throw new Error("DuckDB returned an invalid database name");
            }
            const targetName = `cc_snapshot_${globalThis.crypto.randomUUID().replaceAll("-", "")}`;
            const quotedTarget = quoteIdentifier(targetName);
            await this.#database.registerFileHandle(
              storage.path,
              storage.fileHandle,
              DuckDBDataProtocol.BROWSER_FSACCESS,
              true,
            );
            await this.#database.registerFileHandle(
              `${storage.path}.wal`,
              storage.walHandle,
              DuckDBDataProtocol.BROWSER_FSACCESS,
              true,
            );
            let attached = false;
            try {
              await this.#execute(
                `ATTACH '${storage.path.replaceAll("'", "''")}' AS ${quotedTarget} (READ_WRITE)`,
              );
              attached = true;
              throwIfAborted(signal, "db.export");
              await this.#execute(
                `COPY FROM DATABASE ${quoteIdentifier(sourceName)} TO ${quotedTarget}`,
              );
              throwIfAborted(signal, "db.export");
              await this.#execute(`CHECKPOINT ${quotedTarget}`);
              await this.#execute(`DETACH ${quotedTarget}`);
              attached = false;
              await this.#database.flushFiles();
            } catch (error) {
              if (attached) {
                await this.#execute(`DETACH ${quotedTarget}`).catch(
                  () => undefined,
                );
              }
              throw error;
            }
          },
          release: async (storage) => {
            const failures: unknown[] = [];
            for (const name of [storage.path, `${storage.path}.wal`]) {
              try {
                await this.#database.dropFile(name);
              } catch (error) {
                failures.push(error);
              }
            }
            if (failures.length === 1) throw failures[0];
            if (failures.length > 1) {
              throw new AggregateError(
                failures,
                "DuckDB failed to release browser snapshot files",
              );
            }
          },
        });
      }
      await this.#execute("CHECKPOINT");
      await this.#database.flushFiles();
      return copyBrowserFileToDestination({
        source: this.#fileHandle,
        destination,
        signal,
      });
    });
  }

  async close(): Promise<void> {
    if (this.#closeAttempt !== undefined) return this.#closeAttempt;
    if (this.#closed) return;
    this.#closeRequested = true;
    const closeAttempt = this.#exclusive(async () => {
      const failures: unknown[] = [];
      const attempt = async (
        work: () => Promise<unknown>,
        completed?: (() => void) | undefined,
      ): Promise<void> => {
        try {
          await work();
          completed?.();
        } catch (error) {
          failures.push(error);
        }
      };
      if (!this.#readonly && !this.#checkpointAttempted) {
        this.#checkpointAttempted = true;
        await attempt(() => this.#execute("CHECKPOINT"));
      }
      if (!this.#connectionClosed) {
        await attempt(
          () => this.#connection.close(),
          () => {
            this.#connectionClosed = true;
          },
        );
      }
      if (!this.#filesFlushed) {
        await attempt(
          () => this.#database.flushFiles(),
          () => {
            this.#filesFlushed = true;
          },
        );
      }
      await attempt(
        () => this.#database.terminate(),
        () => {
          this.#closed = true;
        },
      );
      if (failures.length === 1) throw failures[0];
      if (failures.length > 1) {
        throw new AggregateError(
          failures,
          "DuckDB failed while closing the browser database.",
        );
      }
    }, true);
    this.#closeAttempt = closeAttempt;
    try {
      await closeAttempt;
    } finally {
      if (this.#closeAttempt === closeAttempt) this.#closeAttempt = undefined;
    }
  }
}
