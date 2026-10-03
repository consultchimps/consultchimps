import type { ConsultChimpsError } from "@consultchimps/core";

import { databaseError } from "../errors.js";

export const NATIVE_ENGINE_UNAVAILABLE = "DB_NATIVE_ENGINE_UNAVAILABLE";

// Node's codes for a native addon or its loader that is missing or was built
// for another ABI. better-sqlite3 loads its addon in the first constructor
// call, so these can surface there as well as from the module load.
export function isNativeLoadFailure(error: unknown): boolean {
  if (typeof error !== "object" || error === null) return false;
  const code: unknown = Reflect.get(error, "code");
  return code === "MODULE_NOT_FOUND" || code === "ERR_DLOPEN_FAILED";
}

// The native engines load when a database opens, so a missing or mismatched
// binary surfaces there. Name it, rather than letting the open path blame the
// database file.
export function nativeEngineUnavailable(
  format: "sqlite" | "duckdb",
  moduleName: string,
  cause: unknown,
): ConsultChimpsError {
  return databaseError(
    NATIVE_ENGINE_UNAVAILABLE,
    `The ${format === "sqlite" ? "SQLite" : "DuckDB"} engine (${moduleName}) could not be loaded. Reinstall ConsultChimps, or use the portable archive built for this Node.js version, operating system, and architecture.`,
    { format, module: moduleName },
    cause,
  );
}
