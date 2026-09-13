import { throwIfAborted, type RandomAccessFile } from "@consultchimps/core";

import { databaseError } from "./errors.js";

export interface DuckDbSnapshotStorage {
  readonly name: string;
  readonly path: string;
  readonly directory: string;
  copyTo(destination: RandomAccessFile, signal?: AbortSignal): Promise<number>;
  remove(): Promise<void>;
}

export async function exportReadonlyDuckDbSnapshot<
  Storage extends DuckDbSnapshotStorage,
>(options: {
  readonly destination: RandomAccessFile;
  readonly signal?: AbortSignal | undefined;
  allocate(): Promise<Storage>;
  prepare(storage: Storage): Promise<void>;
  release(storage: Storage): Promise<void>;
}): Promise<number> {
  const storage = await options.allocate();
  let bytesWritten: number | undefined;
  let operationFailure: unknown;
  try {
    await options.prepare(storage);
    throwIfAborted(options.signal, "db.browser.export");
    bytesWritten = await storage.copyTo(options.destination, options.signal);
    throwIfAborted(options.signal, "db.browser.export");
  } catch (error) {
    operationFailure = error;
  }

  let releaseFailure: unknown;
  try {
    await options.release(storage);
  } catch (error) {
    releaseFailure = error;
  }
  let removalFailure: unknown;
  if (releaseFailure === undefined) {
    try {
      await storage.remove();
    } catch (error) {
      removalFailure = error;
    }
  }
  if (releaseFailure !== undefined || removalFailure !== undefined) {
    throw databaseError(
      "DB_BROWSER_DUCKDB_SNAPSHOT_CLEANUP_REQUIRED",
      `Browser export could not release its temporary DuckDB snapshot. Close other tabs using this site. Snapshot files may remain in the ${storage.directory} origin-private directory as ${storage.name} and ${storage.name}.wal. If the snapshot opens through BrowserDatabaseRuntime.openDatabase, export it before removing the abandoned files and retrying.`,
      { directory: storage.directory, snapshotName: storage.name },
      new AggregateError(
        [operationFailure, releaseFailure, removalFailure].filter(
          (error) => error !== undefined,
        ),
        "DuckDB browser snapshot cleanup failed",
      ),
    );
  }
  if (operationFailure !== undefined) throw operationFailure;
  if (bytesWritten === undefined) {
    throw new Error("DuckDB snapshot export did not return a byte count");
  }
  return bytesWritten;
}
