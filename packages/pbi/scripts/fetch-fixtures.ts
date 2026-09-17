#!/usr/bin/env node
// Fetch the seven corpus fixtures that are too large to commit.
//
// The two smallest, a-2018-fuzzy.pbix and b-2018-profiling.pbix, are committed
// with their full oracle dumps so the default test run always exercises the
// whole pipeline offline. The other seven are public Microsoft samples pinned
// by repository commit and verified against the SHA-256 digests committed in
// fixtures/oracle-digests.json. Files land in fixtures/fetched/, which is
// ignored by git, and the corpus test reports by name every fixture it skipped
// because its file is not there.
//
//   node scripts/fetch-fixtures.ts          fetch what is missing
//   node scripts/fetch-fixtures.ts --force  refetch everything
import { createHash } from "node:crypto";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, "..");
const target = path.join(root, "fixtures", "fetched");
const force = process.argv.includes("--force");

interface Fixture {
  readonly name: string;
  readonly fileName: string;
  readonly bytes: number;
  readonly sha256: string;
  readonly url: string;
}

const index = JSON.parse(
  await readFile(path.join(root, "fixtures", "oracle-digests.json"), "utf8"),
) as { fixtures: readonly Fixture[] };
const committed = new Set(["a-2018-fuzzy", "b-2018-profiling"]);

await mkdir(target, { recursive: true });
let fetched = 0;
let kept = 0;
let failed = 0;
for (const fixture of index.fixtures) {
  if (committed.has(fixture.name)) continue;
  const destination = path.join(target, fixture.fileName);
  if (!force && existsSync(destination)) {
    const bytes = new Uint8Array(await readFile(destination));
    if (createHash("sha256").update(bytes).digest("hex") === fixture.sha256) {
      kept++;
      continue;
    }
  }
  process.stdout.write(
    `fetching ${fixture.fileName} (${fixture.bytes} bytes)\n`,
  );
  let bytes: Uint8Array;
  try {
    const response = await fetch(fixture.url);
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    bytes = new Uint8Array(await response.arrayBuffer());
  } catch (error) {
    process.stderr.write(`  failed: ${String(error)}\n`);
    failed++;
    continue;
  }
  const digest = createHash("sha256").update(bytes).digest("hex");
  if (digest !== fixture.sha256) {
    process.stderr.write(
      `  digest mismatch: expected ${fixture.sha256}, got ${digest}. Not written.\n`,
    );
    failed++;
    continue;
  }
  await writeFile(destination, bytes);
  fetched++;
}
process.stdout.write(
  `fetched ${fetched}, already present ${kept}, failed ${failed}\n`,
);
process.exitCode = failed > 0 ? 1 : 0;
