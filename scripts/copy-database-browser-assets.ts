import { createRequire } from "node:module";
import { copyFile, mkdir, stat } from "node:fs/promises";
import path from "node:path";

const requireFromDocs = createRequire(
  path.join(import.meta.dirname, "../apps/docs/package.json"),
);
const outputDirectory = path.join(
  import.meta.dirname,
  "../apps/docs/public/database-wasm",
);

// The DuckDB runtime (duckdb-eh.wasm and duckdb-browser-eh.worker.js from
// @duckdb/duckdb-wasm/dist) is left out while the database tool is paused: only
// that tool loads it, and at 35 MB it exceeds the 25 MiB per-file limit of the
// Cloudflare host. Copy both again when the tool returns.
const assets = [
  {
    source: requireFromDocs.resolve("@sqlite.org/sqlite-wasm/sqlite3.wasm"),
    name: "sqlite3.wasm",
  },
] as const;

await mkdir(outputDirectory, { recursive: true });
for (const asset of assets) {
  const destination = path.join(outputDirectory, asset.name);
  await copyFile(asset.source, destination);
  const copied = await stat(destination);
  process.stdout.write(
    `Copied ${asset.name} to the docs app (${copied.size.toLocaleString("en-US")} bytes)\n`,
  );
}
