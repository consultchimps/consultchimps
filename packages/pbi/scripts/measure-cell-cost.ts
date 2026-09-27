#!/usr/bin/env node --expose-gc
// What one cell costs in the JavaScript heap: the decoded value in the reader,
// and the emitted cell in the workbook writer.
//
// The pipeline reserves peakBytes for a column before it decodes it, from the
// catalog's declared row count and the column's declared type, and reserves the
// workbook stage before it builds any worksheet XML. Both need byte figures that
// are measured rather than assumed. A reference is eight bytes; a boxed bigint, a
// boxed double and a dictionary-shared string each cost more than the slot that
// points at them.
//
// The shapes below are the ones `decodeColumn` actually produces: strings come
// from a shared dictionary and repeat, int64 and currency are bigints, double
// and dateTimeSerial are numbers in a mixed-type array, and a null costs a slot.
//
//   node --expose-gc scripts/measure-cell-cost.ts [rows]
//   node --expose-gc scripts/measure-cell-cost.ts --workbook [file ...]
//
// The script imports package sources, so build it first with the bundler this
// repository already depends on:
//
//   node_modules/.bin/esbuild scripts/measure-cell-cost.ts --bundle \
//     --platform=node --format=esm --packages=external \
//     --outfile=scripts/measure-cell-cost.mjs
//
// Run it again after any change to the decoded value representation or to the
// workbook writer.
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import type { PbiValue } from "../src/model.js";
import {
  allocateWorksheets,
  readPbiTables,
  workbookExtent,
  workbookReservationBytes,
} from "../src/pipeline.js";
import { writeWorkbook } from "../src/workbook.js";

const collect = (globalThis as { gc?: () => void }).gc;
if (collect === undefined) {
  process.stderr.write("run with --expose-gc\n");
  process.exit(1);
}

/** Heap and the buffers hanging off it, which is where package bytes live. */
function memoryBytes(): number {
  const usage = process.memoryUsage();
  return usage.heapUsed + usage.external;
}

function settle(): number {
  for (let pass = 0; pass < 4; pass++) collect!();
  return memoryBytes();
}

/** Bytes per cell for one value shape, holding the array alive across the read. */
function measure(
  rows: number,
  shape: string,
  make: (index: number) => PbiValue,
): void {
  settle();
  const before = settle();
  const values: PbiValue[] = new Array<PbiValue>(rows);
  for (let index = 0; index < rows; index++) values[index] = make(index);
  const after = settle();
  const perCell = (after - before) / rows;
  // Reading one value keeps the array reachable past the measurement.
  if (values[rows - 1] === Symbol.iterator) process.stderr.write("");
  process.stdout.write(
    `${shape.padEnd(22)} ${(after - before).toString().padStart(12)} bytes  ${perCell.toFixed(2)} per cell\n`,
  );
}

function measureDecodedValues(rows: number): void {
  // A dictionary column repeats its entries, so the strings themselves are
  // shared and the per-cell cost is the slot plus the amortized entry.
  const dictionary = Array.from(
    { length: 1000 },
    (_, index) => `value ${index} of a shared dictionary`,
  );

  process.stdout.write(`rows per shape: ${rows}\n`);
  measure(rows, "null", () => null);
  measure(rows, "string (shared)", (index) => dictionary[index % 1000]!);
  measure(rows, "string (distinct)", (index) => `row ${index}`);
  measure(rows, "bigint", (index) => BigInt(index));
  measure(
    rows,
    "bigint (past 2^53)",
    (index) => 9_007_199_254_740_993n + BigInt(index),
  );
  measure(rows, "double", (index) => index + 0.5);
  measure(rows, "boolean", (index) => (index & 1) === 1);
}

const fixtureRoot = path.join(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
  "fixtures",
);

/**
 * What the workbook stage costs on a real model, against what it reserves.
 *
 * The stage holds the materialized worksheet XML for every part at once, the
 * package writer's copy of each entry, its compression buffers, the churn the
 * escaping loop produces and the finished package, so the figure is the peak
 * heap the stage reaches rather than what survives it. The XML loop is
 * synchronous and a timer cannot interrupt it, so the peak is the larger of the
 * maximum sampled during compression, where the XML of every sheet is still
 * live, and a reading taken the moment the package is returned.
 *
 * The margin is what the invariant cares about: it must stay above one.
 */
async function measureWorkbookStage(file: string): Promise<void> {
  const input = new Uint8Array(await readFile(file));
  const model = await readPbiTables(input, {
    inputBytes: 256 * 1024 * 1024,
    decodedBytes: 1024 * 1024 * 1024,
    peakBytes: 4 * 1024 * 1024 * 1024,
  });
  const { plans } = allocateWorksheets(model.tables);
  const extent = workbookExtent(plans);
  const reserved = workbookReservationBytes(plans);

  settle();
  const before = settle();
  let peak = before;
  const sampler = setInterval(() => {
    peak = Math.max(peak, memoryBytes());
  }, 1);
  const workbook = await writeWorkbook(plans);
  peak = Math.max(peak, memoryBytes());
  clearInterval(sampler);
  const measured = peak - before;
  process.stdout.write(
    `${path.basename(file).padEnd(24)} ${String(extent.parts).padStart(3)} parts ${String(extent.cells).padStart(10)} cells ${String(extent.rows).padStart(9)} rows ${String(extent.textUnits).padStart(11)} text ${String(workbook.bytes.byteLength).padStart(10)} output ${String(measured).padStart(12)} measured ${String(reserved).padStart(12)} reserved margin ${(reserved / measured).toFixed(2)}\n`,
  );
}

const workbookRequested = process.argv.includes("--workbook");
const positional = process.argv
  .slice(2)
  .filter((argument) => argument !== "--workbook");

if (workbookRequested) {
  const files =
    positional.length > 0
      ? positional
      : [path.join(fixtureRoot, "a-2018-fuzzy.pbix")];
  for (const file of files) await measureWorkbookStage(file);
} else {
  measureDecodedValues(Number(positional[0] ?? 500_000));
}
