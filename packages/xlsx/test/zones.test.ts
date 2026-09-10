/**
 * The time-zone harness testing itself.
 *
 * Every date assertion in this package leans on it, and it works by taking a
 * process-wide value away and putting it back. A helper like that has to be
 * shown to put it back, including when the reading it wrapped threw, and to
 * refuse the one mistake that made it wrong before: two readings in flight at
 * once, each restoring the other's zone.
 */
import { describe, expect, it } from "vitest";

import { forEachZone, inZone, inZoneAsync, ZONES } from "./zones.js";

/** The host zone as the process reports it, however it was left. */
function hostZone(): string | undefined {
  return process.env["TZ"];
}

describe("the time-zone harness", () => {
  it("reads in the zone it was given", () => {
    expect(inZone("UTC", () => new Date().getTimezoneOffset())).toBe(0);
    expect(inZone("Asia/Dubai", () => new Date().getTimezoneOffset())).toBe(
      -240,
    );
  });

  it("puts the host zone back after a reading that returns", () => {
    const before = hostZone();
    inZone("Pacific/Kiritimati", () => "read");
    expect(hostZone()).toBe(before);
  });

  it("puts the host zone back after a reading that throws", () => {
    const before = hostZone();
    expect(() =>
      inZone("Pacific/Kiritimati", () => {
        throw new Error("the reading failed");
      }),
    ).toThrowError(/the reading failed/u);
    expect(hostZone()).toBe(before);
  });

  it("puts the host zone back after an awaited reading that rejects", async () => {
    const before = hostZone();
    await expect(
      inZoneAsync("Pacific/Kiritimati", () => {
        return Promise.reject(new Error("the reading failed"));
      }),
    ).rejects.toThrowError(/the reading failed/u);
    expect(hostZone()).toBe(before);
  });

  it("refuses a second reading while one is in flight", async () => {
    const before = hostZone();
    // The shape the guard exists for: a reading that has not finished, and
    // another started before it did. Whichever one restores first would put
    // the other's zone back as if it were the original.
    await expect(
      inZoneAsync("Asia/Dubai", async () => {
        await Promise.resolve();
        return inZone("UTC", () => "nested");
      }),
    ).rejects.toThrowError(/already set to Asia\/Dubai/u);
    // And the outer reading still put the host zone back on its way out.
    expect(hostZone()).toBe(before);
  });

  it("is usable again after a refusal", () => {
    expect(inZone("UTC", () => new Date().getTimezoneOffset())).toBe(0);
  });

  it("reads every zone in turn, one at a time", async () => {
    const before = hostZone();
    const offsets = await forEachZone(() =>
      Promise.resolve(new Date().getTimezoneOffset()),
    );

    expect(offsets).toHaveLength(ZONES.length);
    expect(offsets[0]).toBe(0);
    expect(new Set(offsets).size).toBe(ZONES.length);
    expect(hostZone()).toBe(before);
  });
});
