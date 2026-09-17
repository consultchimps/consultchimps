#!/usr/bin/env node --expose-gc
// What one decoded cell costs in the JavaScript heap, per value shape.
//
// The pipeline reserves peakBytes for a column before it decodes it, from the
// catalog's declared row count and the column's declared type, so it needs a
// per-type byte figure that is measured rather than assumed. A reference is
// eight bytes; a boxed bigint, a boxed double and a dictionary-shared string
// each cost more than the slot that points at them.
//
// The shapes below are the ones `decodeColumn` actually produces: strings come
// from a shared dictionary and repeat, int64 and currency are bigints, double
// and dateTimeSerial are numbers in a mixed-type array, and a null costs a slot.
//
//   node --expose-gc scripts/measure-cell-cost.ts [rows]
//
// Run it again after any change to the decoded value representation.
import type { PbiValue } from "../src/model.js";

const rows = Number(process.argv[2] ?? 500_000);
const collect = (globalThis as { gc?: () => void }).gc;
if (collect === undefined) {
  process.stderr.write("run with --expose-gc\n");
  process.exit(1);
}

function settle(): number {
  for (let pass = 0; pass < 4; pass++) collect!();
  return process.memoryUsage().heapUsed;
}

/** Bytes per cell for one value shape, holding the array alive across the read. */
function measure(shape: string, make: (index: number) => PbiValue): void {
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

// A dictionary column repeats its entries, so the strings themselves are shared
// and the per-cell cost is the slot plus the amortized entry.
const dictionary = Array.from(
  { length: 1000 },
  (_, index) => `value ${index} of a shared dictionary`,
);

process.stdout.write(`rows per shape: ${rows}\n`);
measure("null", () => null);
measure("string (shared)", (index) => dictionary[index % 1000]!);
measure("string (distinct)", (index) => `row ${index}`);
measure("bigint", (index) => BigInt(index));
measure(
  "bigint (past 2^53)",
  (index) => 9_007_199_254_740_993n + BigInt(index),
);
measure("double", (index) => index + 0.5);
measure("boolean", (index) => (index & 1) === 1);
