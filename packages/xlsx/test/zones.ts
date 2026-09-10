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
 * between two runs is not something an assertion can catch.
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

/** Run one read with the host in a given zone, and put the zone back. */
export function inZone<T>(zone: string, read: () => T): T {
  const original = process.env["TZ"];
  process.env["TZ"] = zone;
  try {
    return read();
  } finally {
    if (original === undefined) {
      delete process.env["TZ"];
    } else {
      process.env["TZ"] = original;
    }
  }
}

/** The same for a read that has to await something. */
export async function inZoneAsync<T>(
  zone: string,
  read: () => Promise<T>,
): Promise<T> {
  const original = process.env["TZ"];
  process.env["TZ"] = zone;
  try {
    return await read();
  } finally {
    if (original === undefined) {
      delete process.env["TZ"];
    } else {
      process.env["TZ"] = original;
    }
  }
}
