import { randomBytes } from "node:crypto";
import {
  appendFileSync,
  mkdirSync,
  readdirSync,
  rmSync,
  statSync,
  writeFileSync,
} from "node:fs";
import os from "node:os";
import path from "node:path";
import { Worker } from "node:worker_threads";
import type { Session } from "node:inspector/promises";

import {
  TASK_LOG_SCHEMA,
  type OperationProgress,
  type OperationResult,
  type TaskLogEnd,
  type TaskLogOutcome,
  type TaskLogRecord,
} from "@consultchimps/core";

/**
 * The CLI's writer for local run records (see TaskLogRecord in core). One
 * JSON Lines file per run, appended synchronously as the run goes, so a run
 * that is interrupted or killed still leaves its progress on disk. Nothing in
 * here may change an operation's output or exit code: every failure disables
 * the record for the rest of the run instead.
 */

export const RUN_RECORD_PATTERN: RegExp =
  /^(\d{8}T\d{9}Z)-[a-z0-9-]+-[0-9a-f]{4}\.(jsonl|cpuprofile)$/u;
const KEEP_RUNS = 100;
const KEEP_DAYS = 30;
const MEMORY_INTERVAL_MS = 1000;

export interface LogEnvironment {
  env: NodeJS.ProcessEnv;
  platform: NodeJS.Platform;
  /** Read only when no environment variable names the folder. */
  homedir: () => string;
}

export function loggingDisabledByEnvironment(env: NodeJS.ProcessEnv): boolean {
  const value = env.CONSULTCHIMPS_LOG?.trim().toLowerCase();
  return (
    value === "off" || value === "0" || value === "false" || value === "no"
  );
}

/** Where run records live: an explicit directory, else the platform's usual place for logs. */
export function resolveLogDirectory(
  context: LogEnvironment = {
    env: process.env,
    platform: process.platform,
    homedir: () => os.homedir(),
  },
): string {
  const { env, platform } = context;
  const homedir = context.homedir;
  if (env.CONSULTCHIMPS_LOG_DIR) return path.resolve(env.CONSULTCHIMPS_LOG_DIR);
  if (platform === "win32") {
    const base = env.LOCALAPPDATA || path.join(homedir(), "AppData", "Local");
    return path.join(base, "consultchimps", "logs");
  }
  if (platform === "darwin") {
    return path.join(homedir(), "Library", "Logs", "consultchimps");
  }
  const state = env.XDG_STATE_HOME || path.join(homedir(), ".local", "state");
  return path.join(state, "consultchimps", "logs");
}

/**
 * Options as a record keeps them: booleans and numbers as given, anything else
 * only as having been set. Option text carries paths, worksheet names, and
 * column names, so it is never kept, even with names.
 */
export function recordedOptions(
  options: Record<string, unknown>,
): Record<string, boolean | number> {
  const recorded: Record<string, boolean | number> = {};
  for (const [key, value] of Object.entries(options)) {
    if (value === undefined) continue;
    recorded[key] =
      typeof value === "boolean" || typeof value === "number" ? value : true;
  }
  return recorded;
}

// Extensions the toolkit reads. Anything else is recorded as "" because the
// part after the last dot of an arbitrary name can be a client's name.
const KNOWN_EXTENSIONS = new Set([
  ".xlsx",
  ".xlsm",
  ".xls",
  ".csv",
  ".pptx",
  ".potx",
  ".pdf",
  ".json",
  ".sqlite",
  ".sqlite3",
  ".db",
  ".duckdb",
  ".ccplan",
]);

export function recordedExtension(inputPath: string): string {
  const extension = path.extname(inputPath).toLowerCase();
  return KNOWN_EXTENSIONS.has(extension) ? extension : "";
}

/** Keep the newest runs, and none older than the age limit; touch only run records. */
export function pruneRunRecords(
  directory: string,
  now: number = Date.now(),
): void {
  const runs = new Map<string, string[]>();
  for (const name of readdirSync(directory)) {
    if (!RUN_RECORD_PATTERN.test(name)) continue;
    const runId = name.slice(0, name.lastIndexOf("."));
    runs.set(runId, [...(runs.get(runId) ?? []), name]);
  }
  const ordered = [...runs.keys()].sort().reverse();
  const oldest = now - KEEP_DAYS * 24 * 60 * 60 * 1000;
  ordered.forEach((runId, position) => {
    const started = runIdTime(runId);
    if (position < KEEP_RUNS && (started === undefined || started >= oldest)) {
      return;
    }
    for (const name of runs.get(runId) ?? []) {
      rmSync(path.join(directory, name), { force: true });
    }
  });
}

export function runIdTime(runId: string): number | undefined {
  const match = /^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})(\d{3})Z/u.exec(
    runId,
  );
  if (!match) return undefined;
  const [, y, mo, d, h, mi, s, ms] = match.map(Number);
  return Date.UTC(y!, mo! - 1, d!, h!, mi!, s!, ms!);
}

function newRunId(command: string, started: Date): string {
  const stamp = started.toISOString().replace(/[-:.]/gu, "");
  const slug = command.toLowerCase().replace(/[^a-z0-9]+/gu, "-") || "run";
  return `${stamp}-${slug}-${randomBytes(2).toString("hex")}`;
}

// Runs on its own thread so it keeps sampling while the main thread is busy in
// a synchronous parse; resident memory belongs to the whole process, so the
// reading is right from here. Heap figures are per thread and are read on the
// main thread instead.
const SAMPLER_SOURCE = `
const { parentPort, workerData } = require("node:worker_threads");
let peak = 0;
const tick = () => {
  const rss = process.memoryUsage.rss();
  if (rss > peak) peak = rss;
  parentPort.postMessage({ t: Date.now() - workerData.start, rss });
};
tick();
const timer = setInterval(tick, workerData.interval);
parentPort.on("message", () => {
  clearInterval(timer);
  tick();
  parentPort.postMessage({ peak });
});
`;

export interface RunRecorderOptions {
  command: string;
  options: Record<string, unknown>;
  names: boolean;
  log: boolean;
  profile: boolean;
  version: string;
  directory?: string;
  /** Called once when recording has to stop; the CLI turns it into one stderr line. */
  onFailure?: (message: string) => void;
}

export interface RunRecorder {
  /** The run's record, when one is being written. */
  readonly recordPath: string | undefined;
  recordInputs(paths: readonly (string | undefined)[]): void;
  progress(progress: OperationProgress): void;
  result(result: OperationResult): void;
  finish(
    outcome: Exclude<TaskLogOutcome, "interrupted">,
    exitCode: number,
    error?: { code: string | null; message: string },
  ): Promise<{ recordPath?: string; profilePath?: string }>;
}

class DisabledRecorder implements RunRecorder {
  readonly recordPath = undefined;
  recordInputs(): void {}
  progress(): void {}
  result(): void {}
  async finish(): Promise<Record<string, never>> {
    return {};
  }
}

export const disabledRecorder: RunRecorder = new DisabledRecorder();

class FileRecorder implements RunRecorder {
  readonly #names: boolean;
  #directory = "";
  readonly #runId: string;
  readonly #start: number;
  readonly #onFailure: ((message: string) => void) | undefined;
  #recordPath: string | undefined;
  #sampler: Worker | undefined;
  #session: Session | undefined;
  #peakRss = 0;
  #peakHeapUsed = 0;
  #result: OperationResult | undefined;
  #finished = false;
  #failed = false;
  #inputCount = 0;
  // Records the interruption, then steps aside: a command with its own
  // cancellation (the db commands) listens too and decides how the run ends.
  // Only when nobody else listens does this end the process, which is what
  // Node would have done without any listener.
  readonly #onSignal = (signal: NodeJS.Signals) => {
    const exitCode = signal === "SIGINT" ? 130 : 143;
    this.#writeEnd({ outcome: "interrupted", exitCode });
    if (process.listenerCount(signal) === 0) process.exit(exitCode);
  };

  constructor(options: RunRecorderOptions) {
    this.#names = options.names;
    const started = new Date();
    this.#start = started.getTime();
    this.#runId = newRunId(options.command, started);
    this.#onFailure = options.onFailure;

    this.#guard(() => {
      this.#directory = options.directory ?? resolveLogDirectory();
      mkdirSync(this.#directory, { recursive: true });
      if (options.log) {
        this.#recordPath = path.join(this.#directory, `${this.#runId}.jsonl`);
        const cpus = os.cpus();
        this.#append({
          type: "run",
          schema: TASK_LOG_SCHEMA,
          runId: this.#runId,
          startedAt: started.toISOString(),
          command: options.command,
          surface: "cli",
          version: options.version,
          runtime: `node ${process.versions.node}`,
          platform: process.platform,
          release: os.release(),
          arch: process.arch,
          cpus: cpus.length,
          cpuModel: cpus[0]?.model.trim() ?? "",
          totalMemory: os.totalmem(),
          names: this.#names,
          options: recordedOptions(options.options),
        });
        this.#startSampler();
        process.once("SIGINT", this.#onSignal);
        process.once("SIGTERM", this.#onSignal);
      }
    });
  }

  get recordPath(): string | undefined {
    return this.#recordPath;
  }

  async startProfile(): Promise<void> {
    if (this.#failed) return;
    try {
      // Loaded only when asked for, so a runtime without the inspector still
      // runs every command.
      const { Session: InspectorSession } =
        await import("node:inspector/promises");
      const session = new InspectorSession();
      session.connect();
      await session.post("Profiler.enable");
      await session.post("Profiler.start");
      this.#session = session;
    } catch (error) {
      this.#fail(error);
    }
  }

  recordInputs(paths: readonly (string | undefined)[]): void {
    this.#guard(() => {
      for (const inputPath of paths) {
        if (typeof inputPath !== "string") continue;
        this.#inputCount += 1;
        let bytes = -1;
        try {
          bytes = statSync(inputPath).size;
        } catch {
          // A missing input is the operation's to report; the record notes it
          // with a size of -1 rather than failing first.
        }
        this.#append({
          type: "input",
          index: this.#inputCount,
          extension: recordedExtension(inputPath),
          bytes,
          ...(this.#names ? { name: path.basename(inputPath) } : {}),
        });
      }
    });
  }

  progress(progress: OperationProgress): void {
    this.#guard(() => {
      this.#sampleHeap();
      this.#append({
        type: "progress",
        t: Date.now() - this.#start,
        operation: progress.operation,
        stage: progress.stage,
        completed: progress.completed,
        total: progress.total,
        ...(progress.measures ? { measures: progress.measures } : {}),
        ...(this.#names && progress.detail !== undefined
          ? { detail: progress.detail }
          : {}),
      });
    });
  }

  result(result: OperationResult): void {
    this.#result = result;
  }

  async finish(
    outcome: Exclude<TaskLogOutcome, "interrupted">,
    exitCode: number,
    error?: { code: string | null; message: string },
  ): Promise<{ recordPath?: string; profilePath?: string }> {
    process.removeListener("SIGINT", this.#onSignal);
    process.removeListener("SIGTERM", this.#onSignal);
    const profilePath = await this.#stopProfile();
    await this.#stopSampler();
    const end: Omit<TaskLogEnd, "type" | "t" | "peakRss" | "peakHeapUsed"> = {
      outcome,
      exitCode,
    };
    if (this.#result) {
      end.operation = this.#result.operation;
      end.metrics = this.#result.metrics;
      end.warnings = this.#result.warnings.length;
      end.artifacts = this.#result.artifacts.length;
    }
    if (error) {
      end.error = this.#names
        ? { code: error.code, message: error.message }
        : { code: error.code };
    }
    if (profilePath) end.profile = path.basename(profilePath);
    this.#writeEnd(end);
    // Pruning is housekeeping: the record is already written, so a failure
    // here is not reported as the record stopping.
    try {
      if (this.#directory) pruneRunRecords(this.#directory);
    } catch {
      // Left for the next run to retry.
    }
    const paths: { recordPath?: string; profilePath?: string } = {};
    if (this.#recordPath) paths.recordPath = this.#recordPath;
    if (profilePath) paths.profilePath = profilePath;
    return paths;
  }

  #writeEnd(
    end: Omit<TaskLogEnd, "type" | "t" | "peakRss" | "peakHeapUsed">,
  ): void {
    if (this.#finished) return;
    this.#finished = true;
    this.#guard(() => {
      this.#sampleHeap();
      this.#peakRss = Math.max(this.#peakRss, process.memoryUsage.rss());
      this.#append({
        type: "end",
        t: Date.now() - this.#start,
        peakRss: this.#peakRss,
        peakHeapUsed: this.#peakHeapUsed,
        ...end,
      });
    });
  }

  #startSampler(): void {
    const sampler = new Worker(SAMPLER_SOURCE, {
      eval: true,
      workerData: { start: this.#start, interval: MEMORY_INTERVAL_MS },
    });
    sampler.unref();
    sampler.on(
      "message",
      (message: { t?: number; rss?: number; peak?: number }) => {
        if (message.peak !== undefined) {
          this.#peakRss = Math.max(this.#peakRss, message.peak);
          return;
        }
        if (message.t === undefined || message.rss === undefined) return;
        this.#peakRss = Math.max(this.#peakRss, message.rss);
        this.#guard(() => {
          const { heapUsed, external } = this.#sampleHeap();
          this.#append({
            type: "memory",
            t: message.t!,
            rss: message.rss!,
            heapUsed,
            external,
          });
        });
      },
    );
    sampler.on("error", (error) => this.#fail(error));
    this.#sampler = sampler;
  }

  async #stopSampler(): Promise<void> {
    const sampler = this.#sampler;
    if (!sampler) return;
    this.#sampler = undefined;
    await new Promise<void>((resolve) => {
      const done = () => resolve();
      const timer = setTimeout(done, 1000);
      sampler.on("message", (message: { peak?: number }) => {
        if (message.peak !== undefined) {
          clearTimeout(timer);
          done();
        }
      });
      sampler.postMessage("stop");
    });
    await sampler.terminate();
  }

  async #stopProfile(): Promise<string | undefined> {
    const session = this.#session;
    if (!session) return undefined;
    this.#session = undefined;
    try {
      const { profile } = await session.post("Profiler.stop");
      session.disconnect();
      const profilePath = path.join(
        this.#directory,
        `${this.#runId}.cpuprofile`,
      );
      writeFileSync(profilePath, JSON.stringify(profile));
      return profilePath;
    } catch (error) {
      this.#fail(error);
      return undefined;
    }
  }

  #sampleHeap(): { heapUsed: number; external: number } {
    const { heapUsed, external } = process.memoryUsage();
    this.#peakHeapUsed = Math.max(this.#peakHeapUsed, heapUsed);
    return { heapUsed, external };
  }

  #append(record: TaskLogRecord): void {
    // The end line is the last line: once it is written, a late memory sample
    // (a command that handles its own cancellation keeps running after an
    // interruption was recorded) has nowhere to go.
    if (!this.#recordPath || (this.#finished && record.type !== "end")) return;
    appendFileSync(this.#recordPath, `${JSON.stringify(record)}\n`);
  }

  #guard(action: () => void): void {
    try {
      action();
    } catch (error) {
      this.#fail(error);
    }
  }

  #fail(error: unknown): void {
    if (this.#failed) return;
    this.#failed = true;
    this.#recordPath = undefined;
    const sampler = this.#sampler;
    this.#sampler = undefined;
    void sampler?.terminate();
    this.#onFailure?.(error instanceof Error ? error.message : String(error));
  }
}

let active: RunRecorder = disabledRecorder;

/** The recorder for the command now running, or one that records nothing. */
export function currentRecorder(): RunRecorder {
  return active;
}

export async function startRunRecorder(
  options: RunRecorderOptions,
): Promise<RunRecorder> {
  if (!options.log && !options.profile) {
    active = disabledRecorder;
    return active;
  }
  const recorder = new FileRecorder(options);
  if (options.profile) await recorder.startProfile();
  active = recorder;
  return recorder;
}
