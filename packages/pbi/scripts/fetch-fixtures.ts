#!/usr/bin/env node
// Fetch the eight corpus fixtures the repository does not commit.
//
// One fixture, a-2018-fuzzy.pbix, is committed with its full oracle dump so the
// default test run always exercises the whole pipeline offline. Its cells are
// first names and two-letter state codes. Every other sample in the corpus
// carries person-like names, street addresses or telephone numbers, which the
// repository does not commit even from a public sample. Those eight are public
// Microsoft samples pinned by repository commit and verified against the
// SHA-256 digests in fixtures/oracle-digests.json. They land in
// fixtures/fetched/, which git ignores. Locally the corpus test names every
// fixture it skipped; in CI it fetches a missing one itself, which the pinned
// digests make safe.
//
//   node scripts/fetch-fixtures.ts          fetch what is missing
//   node scripts/fetch-fixtures.ts --force  refetch everything
import { createHash } from "node:crypto";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

export interface Fixture {
  readonly name: string;
  readonly fileName: string;
  readonly bytes: number;
  readonly sha256: string;
  readonly url: string;
}

/** The one fixture whose cells the repository is willing to carry. */
export const COMMITTED_FIXTURES: ReadonlySet<string> = new Set([
  "a-2018-fuzzy",
]);

/**
 * Fetch one pinned fixture and verify its digest before anything is written, so
 * a mismatched or truncated download never lands on disk. Exported so the
 * corpus test can fetch a missing fixture itself when it runs in CI.
 */
export async function fetchFixture(
  fixture: Fixture,
  destination: string,
): Promise<void> {
  let bytes: Uint8Array;
  try {
    const response = await fetch(fixture.url);
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    bytes = new Uint8Array(await response.arrayBuffer());
  } catch (error) {
    throw new Error(`could not fetch ${fixture.fileName}`, { cause: error });
  }
  const digest = createHash("sha256").update(bytes).digest("hex");
  if (digest !== fixture.sha256)
    throw new Error(
      `digest mismatch for ${fixture.fileName}: expected ${fixture.sha256}, got ${digest}`,
    );
  await mkdir(path.dirname(destination), { recursive: true });
  await writeFile(destination, bytes);
}

// Everything below runs only when this file is executed as a script. Importing
// it for `fetchFixture`, as the corpus test does, must not start a download.
const invoked = process.argv[1];
if (
  invoked !== undefined &&
  path.resolve(fileURLToPath(import.meta.url)) === path.resolve(invoked)
) {
  const here = path.dirname(fileURLToPath(import.meta.url));
  const root = path.resolve(here, "..");
  const target = path.join(root, "fixtures", "fetched");
  const force = process.argv.includes("--force");
  const index = JSON.parse(
    await readFile(path.join(root, "fixtures", "oracle-digests.json"), "utf8"),
  ) as { fixtures: readonly Fixture[] };

  await mkdir(target, { recursive: true });
  let fetched = 0;
  let kept = 0;
  let failed = 0;
  for (const fixture of index.fixtures) {
    if (COMMITTED_FIXTURES.has(fixture.name)) continue;
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
    try {
      await fetchFixture(fixture, destination);
      fetched++;
    } catch (error) {
      process.stderr.write(`  ${String(error)}\n`);
      failed++;
    }
  }
  process.stdout.write(
    `fetched ${fetched}, already present ${kept}, failed ${failed}\n`,
  );
  process.exitCode = failed > 0 ? 1 : 0;
}
