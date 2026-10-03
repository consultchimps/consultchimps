import type { ConsultChimpsError } from "@consultchimps/core";

import { databaseError } from "../errors.js";

export const NATIVE_ENGINE_UNAVAILABLE = "DB_NATIVE_ENGINE_UNAVAILABLE";

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
