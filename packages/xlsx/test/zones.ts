/**
 * Reading the same workbook with the host in different time zones.
 *
 * A value read out of a workbook is a function of what the workbook holds, and
 * of nothing about the machine reading it. Dates are where that is easy to
 * lose: a `Date` has two faces, and composing or reading the local one makes
 * the text a function of the host offset as well, so the same cell becomes a
 * different calendar day for a reader in Dubai than for one in California.
 *
 * Node applies `process.env.TZ` on assignment, so one process can read a value
 * from every zone below and compare the answers inside a single assertion.
 * Spawning the file once per zone would test the same thing, but a difference
 * between two runs is not something an assertion can catch, so the comparison
 * would have to move outside the suite.
 *
 * The zone is process-global, which makes this state two readers must never
 * hold at once. Vitest runs each test file in its own worker and runs the tests
 * in a file one after another, so the only way to overlap two readings is to
 * start them without awaiting - which `firstRowPerZone` did, with `Promise.all`,
 * so each call captured another call's zone as the zone to restore and could
 * leave the worker in a test zone. The guard below turns that mistake into a
 * loud failure instead of a quiet one, and `forEachZone` is the sequential way
 * to ask the question that made it tempting.
 */

/**
 * UTC, east of it by a whole hour and by a half hour, west of it, and the
 * furthest east there is.
 */
export const ZONES = [
  "UTC",
  "Asia/Dubai",
  "Asia/Kolkata",
  "America/Los_Angeles",
  "Pacific/Kiritimati",
] as const;

/** The zone a reading is running in, or undefined when none is. */
let reading: string | undefined;

/**
 * Take the host zone, returning what it was so the exit can put it back.
 *
 * Refuses to be entered while another reading holds the zone: the value it
 * would capture is that reading's zone rather than the true original, and
 * restoring it would leave the worker in a test zone for everything after.
 */
function enterZone(zone: string): string | undefined {
  if (reading !== undefined) {
    throw new Error(
      `The host time zone is already set to ${reading} by a reading that has not finished. Time zone readings cannot overlap, because the zone is one process-wide value; run them one after another, with forEachZone or an awaited loop.`,
    );
  }
  const original = process.env["TZ"];
  reading = zone;
  process.env["TZ"] = zone;
  return original;
}

/** Put the host zone back exactly as it was, including having been unset. */
function leaveZone(original: string | undefined): void {
  reading = undefined;
  if (original === undefined) {
    delete process.env["TZ"];
  } else {
    process.env["TZ"] = original;
  }
}

/** Run one read with the host in a given zone, and put the zone back. */
export function inZone<T>(zone: string, read: () => T): T {
  const original = enterZone(zone);
  try {
    return read();
  } finally {
    leaveZone(original);
  }
}

/** The same for a read that has to await something. */
export async function inZoneAsync<T>(
  zone: string,
  read: () => Promise<T>,
): Promise<T> {
  const original = enterZone(zone);
  try {
    return await read();
  } finally {
    leaveZone(original);
  }
}

/**
 * Read the same thing from every zone, one zone at a time.
 *
 * Sequential on purpose: each reading holds the process-wide zone for as long
 * as it takes, including across every await inside it, so two of them cannot
 * be in flight together.
 */
export async function forEachZone<T>(
  read: (zone: string) => Promise<T>,
): Promise<T[]> {
  const results: T[] = [];
  for (const zone of ZONES) {
    // Awaited in the loop on purpose: one reading at a time is the invariant.
    results.push(await inZoneAsync(zone, () => read(zone)));
  }
  return results;
}
