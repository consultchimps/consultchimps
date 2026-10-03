import { createRequire } from "node:module";

import type BetterSqlite3 from "better-sqlite3";

// Loaded on first open rather than imported, so the single-file portable CLI
// bundle does not load the native library for commands that open no database.
// A static import would be hoisted to the top of that bundle.
export function betterSqlite3(): typeof BetterSqlite3 {
  return createRequire(import.meta.url)(
    "better-sqlite3",
  ) as typeof BetterSqlite3;
}
