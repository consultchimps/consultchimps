import { describe, expect, it } from "vitest";

import { summarizeTaskLog, type TaskLogRecord } from "../src/index.js";

const run: TaskLogRecord = {
  type: "run",
  schema: 1,
  runId: "20260929T100000000Z-sheets-consolidate-abcd",
  startedAt: "2026-09-29T10:00:00.000Z",
  command: "sheets consolidate",
  surface: "cli",
  version: "0.12.0",
  runtime: "node 24.0.0",
  platform: "linux",
  release: "6.0",
  arch: "x64",
  cpus: 8,
  cpuModel: "test",
  totalMemory: 1,
  names: false,
  options: {},
};

describe("summarizeTaskLog", () => {
  it("charges each stage and unit the time since the previous event", () => {
    const summary = summarizeTaskLog([
      run,
      { type: "input", index: 1, extension: ".xlsx", bytes: 10 },
      { type: "input", index: 2, extension: ".xlsx", bytes: 20 },
      {
        type: "progress",
        t: 100,
        operation: "sheets.consolidate",
        stage: "reading-workbooks",
        completed: 1,
        total: 2,
        measures: { rows: 5 },
      },
      {
        type: "progress",
        t: 400,
        operation: "sheets.consolidate",
        stage: "reading-workbooks",
        completed: 2,
        total: 2,
      },
      { type: "memory", t: 450, rss: 900, heapUsed: 300, external: 1 },
      {
        type: "progress",
        t: 1000,
        operation: "sheets.consolidate",
        stage: "writing-output",
        completed: 1,
        total: 1,
      },
      {
        type: "end",
        t: 1010,
        outcome: "ok",
        exitCode: 0,
        peakRss: 800,
        peakHeapUsed: 400,
      },
    ]);

    expect(summary.outcome).toBe("ok");
    expect(summary.durationMs).toBe(1010);
    expect(summary.peakRss).toBe(900);
    expect(summary.peakHeapUsed).toBe(400);
    expect(summary.stages).toEqual([
      {
        operation: "sheets.consolidate",
        stage: "reading-workbooks",
        durationMs: 400,
        units: 2,
      },
      {
        operation: "sheets.consolidate",
        stage: "writing-output",
        durationMs: 600,
        units: 1,
      },
    ]);
    expect(summary.slowestUnits.map((unit) => unit.durationMs)).toEqual([
      600, 300, 100,
    ]);
    const second = summary.slowestUnits[1]!;
    expect(second.input?.bytes).toBe(20);
    expect(summary.slowestUnits[2]!.measures).toEqual({ rows: 5 });
    // A stage whose total is not the input count is not matched to inputs.
    expect(summary.slowestUnits[0]!.input).toBeUndefined();
  });

  it("reports a record without an end line as running", () => {
    const summary = summarizeTaskLog([
      run,
      {
        type: "progress",
        t: 50,
        operation: "pdf.split",
        stage: "writing-pages",
        completed: 1,
        total: 3,
      },
    ]);
    expect(summary.outcome).toBe("running");
    expect(summary.durationMs).toBe(50);
    expect(summary.end).toBeUndefined();
  });
});
