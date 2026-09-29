/**
 * The local record of one run: what ran, on what kind of machine, how long each
 * stage and unit took, how much memory it used, and how it ended. A surface
 * writes the lines in this order as the run goes, so an interrupted run still
 * leaves everything up to the interruption.
 *
 * A record holds counts, sizes, and timings. Names (file names, progress
 * details, error messages) appear only when the surface was asked to keep them
 * and `TaskLogRun.names` says so.
 */
export const TASK_LOG_SCHEMA = 1;

export interface TaskLogRun {
  type: "run";
  schema: typeof TASK_LOG_SCHEMA;
  runId: string;
  startedAt: string;
  command: string;
  surface: string;
  version: string;
  runtime: string;
  platform: string;
  release: string;
  arch: string;
  cpus: number;
  cpuModel: string;
  totalMemory: number;
  names: boolean;
  /** Options set for the run: booleans and numbers as given, other values as `true`, unless names are kept. */
  options: Record<string, boolean | number | string>;
}

export interface TaskLogInput {
  type: "input";
  /** 1-based position among the run's inputs. */
  index: number;
  extension: string;
  bytes: number;
  name?: string;
}

export interface TaskLogProgress {
  type: "progress";
  /** Milliseconds since the run started. */
  t: number;
  operation: string;
  stage: string;
  completed: number;
  total: number;
  measures?: Record<string, number>;
  detail?: string;
}

export interface TaskLogMemory {
  type: "memory";
  t: number;
  rss: number;
  heapUsed: number;
  external: number;
}

export type TaskLogOutcome = "ok" | "error" | "interrupted";

export interface TaskLogEnd {
  type: "end";
  t: number;
  outcome: TaskLogOutcome;
  exitCode: number;
  peakRss: number;
  peakHeapUsed: number;
  operation?: string;
  metrics?: Record<string, number>;
  warnings?: number;
  artifacts?: number;
  error?: { code: string | null; message?: string };
  profile?: string;
}

export type TaskLogRecord =
  TaskLogRun | TaskLogInput | TaskLogProgress | TaskLogMemory | TaskLogEnd;

export interface TaskLogStageSummary {
  operation: string;
  stage: string;
  durationMs: number;
  units: number;
}

export interface TaskLogUnitSummary {
  stage: string;
  completed: number;
  total: number;
  durationMs: number;
  measures?: Record<string, number>;
  /** The input this unit read, when the stage counts one unit per input. */
  input?: TaskLogInput;
  detail?: string;
}

export interface TaskLogSummary {
  run?: TaskLogRun;
  inputs: TaskLogInput[];
  /** `running` when the record has no end line: still going, or killed. */
  outcome: TaskLogOutcome | "running";
  durationMs: number;
  peakRss: number;
  peakHeapUsed: number;
  stages: TaskLogStageSummary[];
  /** Every unit, slowest first. */
  slowestUnits: TaskLogUnitSummary[];
  end?: TaskLogEnd;
}

/**
 * Reduce a run's lines to durations. A stage lasts from the event before its
 * first one to its last one, and a unit from the event before it to its own, so
 * time spent between events is charged to the unit that finished. A unit is
 * matched to an input when its stage counts exactly one unit per input.
 */
export function summarizeTaskLog(
  records: readonly TaskLogRecord[],
): TaskLogSummary {
  let run: TaskLogRun | undefined;
  let end: TaskLogEnd | undefined;
  const inputs: TaskLogInput[] = [];
  const progress: TaskLogProgress[] = [];
  let lastT = 0;
  let peakRss = 0;
  let peakHeapUsed = 0;

  for (const record of records) {
    switch (record.type) {
      case "run":
        run = record;
        break;
      case "input":
        inputs.push(record);
        break;
      case "progress":
        progress.push(record);
        lastT = Math.max(lastT, record.t);
        break;
      case "memory":
        peakRss = Math.max(peakRss, record.rss);
        peakHeapUsed = Math.max(peakHeapUsed, record.heapUsed);
        lastT = Math.max(lastT, record.t);
        break;
      case "end":
        end = record;
        peakRss = Math.max(peakRss, record.peakRss);
        peakHeapUsed = Math.max(peakHeapUsed, record.peakHeapUsed);
        lastT = Math.max(lastT, record.t);
        break;
    }
  }

  const stages: TaskLogStageSummary[] = [];
  const units: TaskLogUnitSummary[] = [];
  let previousT = 0;
  for (const event of progress) {
    const durationMs = Math.max(0, event.t - previousT);
    previousT = event.t;
    const current = stages.at(-1);
    if (
      current !== undefined &&
      current.operation === event.operation &&
      current.stage === event.stage
    ) {
      current.durationMs += durationMs;
      current.units += 1;
    } else {
      stages.push({
        operation: event.operation,
        stage: event.stage,
        durationMs,
        units: 1,
      });
    }

    const unit: TaskLogUnitSummary = {
      stage: event.stage,
      completed: event.completed,
      total: event.total,
      durationMs,
    };
    if (event.measures !== undefined) unit.measures = event.measures;
    if (event.detail !== undefined) unit.detail = event.detail;
    const input =
      inputs.length > 0 && event.total === inputs.length
        ? inputs.find((candidate) => candidate.index === event.completed)
        : undefined;
    if (input !== undefined) unit.input = input;
    units.push(unit);
  }

  const summary: TaskLogSummary = {
    inputs,
    outcome: end?.outcome ?? "running",
    durationMs: end?.t ?? lastT,
    peakRss,
    peakHeapUsed,
    stages,
    slowestUnits: units.sort((a, b) => b.durationMs - a.durationMs),
  };
  if (run !== undefined) summary.run = run;
  if (end !== undefined) summary.end = end;
  return summary;
}
