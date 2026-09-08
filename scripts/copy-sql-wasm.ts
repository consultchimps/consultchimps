import { createRequire } from "node:module";
import { copyFileSync, mkdirSync, readFileSync, statSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

// Copy the sql.js WebAssembly binary into the docs app's public directory so the
// static export serves it from our own origin, never a CDN. This is a
// load-bearing offline and local-first rule (docs/adr/0003 Decision 2): the
// workspace worker fetches the wasm through `locateFile` at `<basePath>/sql-wasm/`,
// and that path only resolves if the file sits under apps/docs/public/sql-wasm/.
//
// The copy is derived from the sql.js version this workspace has installed
// rather than committed, so it can never drift from the pinned dependency: the
// docs build regenerates it every time. sql.js is resolved through @consultchimps/db,
// which is the package that actually depends on it.

const workspaceRoot = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
);

// Anchor resolution at the db package, whose manifest declares sql.js, so this
// works regardless of the package manager's hoisting layout.
const dbManifest = path.join(workspaceRoot, "packages", "db", "package.json");
const requireFromDb = createRequire(dbManifest);

// The production wasm binaries sql.js ships. Which one is requested depends on
// the glue the bundler picks: Node builds ask for `sql-wasm.wasm`, and the
// browser build (chosen through the package's `browser` field, which is what
// Next bundles into the worker) asks for `sql-wasm-browser.wasm`. Both are
// copied so `locateFile` resolves whichever the active build names. The debug
// variants are left out: only debug glue asks for them, and it is never bundled
// in production.
const wasmFiles = ["sql-wasm.wasm", "sql-wasm-browser.wasm"] as const;

const destinationDirectory = path.join(
  workspaceRoot,
  "apps",
  "docs",
  "public",
  "sql-wasm",
);

mkdirSync(destinationDirectory, { recursive: true });

for (const fileName of wasmFiles) {
  const sourceWasm = requireFromDb.resolve(`sql.js/dist/${fileName}`);
  const destinationWasm = path.join(destinationDirectory, fileName);
  const sourceSize = statSync(sourceWasm).size;

  // Skip the write when the destination already matches, so an incremental
  // build does not rewrite an identical file. A missing or unreadable
  // destination simply counts as out of date.
  const alreadyCurrent = ((): boolean => {
    try {
      return readFileSync(destinationWasm).equals(readFileSync(sourceWasm));
    } catch {
      return false;
    }
  })();

  if (alreadyCurrent) {
    process.stdout.write(
      `sql.js wasm already up to date at apps/docs/public/sql-wasm/${fileName} (${sourceSize} bytes).\n`,
    );
  } else {
    copyFileSync(sourceWasm, destinationWasm);
    process.stdout.write(
      `Copied sql.js wasm to apps/docs/public/sql-wasm/${fileName} (${sourceSize} bytes).\n`,
    );
  }
}
