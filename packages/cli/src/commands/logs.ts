import { existsSync, readdirSync, readFileSync } from "node:fs";
import path from "node:path";

import {
  ConsultChimpsError,
  summarizeTaskLog,
  type TaskLogRecord,
  type TaskLogRun,
  type TaskLogSummary,
  type TaskLogUnitSummary,
} from "@consultchimps/core";
import { InvalidArgumentError, type Command } from "commander";

import { RUN_RECORD_PATTERN, resolveLogDirectory } from "../task-log.js";
import { withoutTerminalControls } from "../text.js";

export interface LogsCommandOutput {
  data(value: unknown, humanText: string): void;
}

interface RunListing {
  runId: string;
  recordPath: string;
  summary: TaskLogSummary;
}

/** A run's lines, skipping any that do not parse: a killed run can leave half a line. */
export function readRunRecord(recordPath: string): TaskLogRecord[] {
  const records: TaskLogRecord[] = [];
  for (const line of readFileSync(recordPath, "utf8").split("\n")) {
    if (line.trim() === "") continue;
    try {
      records.push(JSON.parse(line) as TaskLogRecord);
    } catch {
      // Ignored: a partial trailing line from an interrupted write.
    }
  }
  return records;
}

function listRuns(directory: string): RunListing[] {
  if (!existsSync(directory)) return [];
  return readdirSync(directory)
    .filter((name) => RUN_RECORD_PATTERN.test(name) && name.endsWith(".jsonl"))
    .sort()
    .reverse()
    .map((name) => {
      const recordPath = path.join(directory, name);
      return {
        runId: name.slice(0, -".jsonl".length),
        recordPath,
        summary: summarizeTaskLog(readRunRecord(recordPath)),
      };
    });
}

export function formatDuration(ms: number): string {
  if (ms < 1000) return `${Math.round(ms)}ms`;
  const seconds = ms / 1000;
  if (seconds < 60) return `${seconds.toFixed(1)}s`;
  const minutes = Math.floor(seconds / 60);
  const rest = Math.round(seconds - minutes * 60);
  return `${minutes}m ${String(rest).padStart(2, "0")}s`;
}

export function formatBytes(bytes: number): string {
  if (bytes < 0) return "unknown size";
  if (bytes < 1024) return `${bytes} B`;
  const units = ["KB", "MB", "GB", "TB"];
  let value = bytes / 1024;
  let unit = 0;
  while (value >= 1024 && unit < units.length - 1) {
    value /= 1024;
    unit += 1;
  }
  return `${value.toFixed(value < 10 ? 2 : 1)} ${units[unit]}`;
}

function when(run: TaskLogRun | undefined): string {
  return run ? run.startedAt.replace("T", " ").slice(0, 19) : "unknown";
}

function describeUnit(unit: TaskLogUnitSummary): string {
  const parts = [`${unit.stage} ${unit.completed}/${unit.total}`];
  const profile: string[] = [];
  if (unit.input) {
    profile.push(
      unit.input.name !== undefined
        ? withoutTerminalControls(unit.input.name)
        : `input ${unit.input.index}`,
      `${unit.input.extension || "no extension"} ${formatBytes(unit.input.bytes)}`,
    );
  } else if (unit.detail !== undefined) {
    profile.push(withoutTerminalControls(unit.detail));
  }
  for (const [key, value] of Object.entries(unit.measures ?? {})) {
    profile.push(`${key} ${value.toLocaleString("en-US")}`);
  }
  if (profile.length > 0) parts.push(`(${profile.join(", ")})`);
  return parts.join(" ");
}

function formatList(runs: readonly RunListing[], directory: string): string {
  if (runs.length === 0) {
    return `No run records yet in ${withoutTerminalControls(directory)}\n`;
  }
  const rows = runs.map((listing) => [
    when(listing.summary.run),
    listing.summary.run?.command ?? "unknown",
    listing.summary.outcome,
    formatDuration(listing.summary.durationMs),
    formatBytes(listing.summary.peakRss),
    listing.runId,
  ]);
  const header = [
    "Started (UTC)",
    "Command",
    "Outcome",
    "Duration",
    "Peak memory",
    "Run",
  ];
  const widths = header.map((title, column) =>
    Math.max(title.length, ...rows.map((row) => row[column]!.length)),
  );
  const line = (cells: string[]) =>
    cells
      .map((cell, column) => cell.padEnd(widths[column]!))
      .join("  ")
      .trimEnd();
  return `${[line(header), ...rows.map(line)].join("\n")}\n`;
}

function formatShow(listing: RunListing): string {
  const { summary } = listing;
  const run = summary.run;
  const lines = [
    `${run?.command ?? "unknown command"}: ${summary.outcome}, ${formatDuration(summary.durationMs)}, peak memory ${formatBytes(summary.peakRss)}`,
  ];
  if (run) {
    lines.push(
      `Started ${when(run)} UTC on consultchimps ${run.version}, ${run.runtime}, ${run.platform} ${run.arch}, ${run.cpus} CPUs, ${formatBytes(run.totalMemory)}`,
    );
  }
  lines.push(`Record: ${withoutTerminalControls(listing.recordPath)}`);
  if (summary.end?.profile) {
    lines.push(
      `CPU profile: ${withoutTerminalControls(path.join(path.dirname(listing.recordPath), summary.end.profile))}`,
    );
  }
  if (summary.inputs.length > 0) {
    const total = summary.inputs.reduce(
      (sum, input) => sum + Math.max(0, input.bytes),
      0,
    );
    lines.push(
      `Inputs: ${summary.inputs.length}, ${formatBytes(total)} in total`,
    );
  }
  if (summary.stages.length > 0) {
    lines.push("", "Stages:");
    for (const stage of summary.stages) {
      lines.push(
        `  ${stage.stage}: ${formatDuration(stage.durationMs)} over ${stage.units} ${stage.units === 1 ? "step" : "steps"}`,
      );
    }
    lines.push("", "Slowest steps:");
    for (const unit of summary.slowestUnits.slice(0, 5)) {
      lines.push(`  ${formatDuration(unit.durationMs)}  ${describeUnit(unit)}`);
    }
  }
  const metrics = Object.entries(summary.end?.metrics ?? {});
  if (metrics.length > 0) {
    lines.push(
      "",
      `Metrics: ${metrics.map(([key, value]) => `${key} ${value.toLocaleString("en-US")}`).join(", ")}`,
    );
  }
  if (summary.end?.error) {
    const { code, message } = summary.end.error;
    lines.push(
      "",
      `Error: ${code ?? "unexpected"}${message !== undefined ? `: ${withoutTerminalControls(message)}` : ""}`,
    );
  }
  if (summary.outcome === "running") {
    lines.push(
      "",
      "No end line: the run is still going, or it was stopped before it could write one.",
    );
  }
  return `${lines.join("\n")}\n`;
}

export function registerLogsCommands(
  program: Command,
  output: LogsCommandOutput,
): void {
  const logs = program
    .command("logs")
    .description("list and read the local records of earlier runs")
    .addHelpText(
      "after",
      "\nExamples:\n  consultchimps logs\n  consultchimps logs show\n  consultchimps logs path\n\nRecords stay on this machine. They hold counts, sizes, and timings, and names only for runs made with --log-names.\n",
    );

  logs
    .command("list", { isDefault: true })
    .description("list recent runs, newest first")
    .option(
      "--limit <count>",
      "how many runs to list",
      (value) => {
        const count = Number(value);
        if (!Number.isInteger(count) || count < 1) {
          throw new InvalidArgumentError("Expected a positive integer.");
        }
        return count;
      },
      20,
    )
    .action((options: { limit: number }) => {
      const directory = resolveLogDirectory();
      const runs = listRuns(directory).slice(0, options.limit);
      output.data(
        {
          directory,
          runs: runs.map((listing) => ({
            runId: listing.runId,
            recordPath: listing.recordPath,
            command: listing.summary.run?.command ?? null,
            startedAt: listing.summary.run?.startedAt ?? null,
            outcome: listing.summary.outcome,
            durationMs: listing.summary.durationMs,
            peakRss: listing.summary.peakRss,
          })),
        },
        formatList(runs, directory),
      );
    });

  logs
    .command("show")
    .description(
      "summarise one run: stages, slowest steps, memory, and outcome",
    )
    .argument("[run]", "a run id or its start, or latest", "latest")
    .action((run: string) => {
      const directory = resolveLogDirectory();
      const runs = listRuns(directory);
      const listing =
        run === "latest"
          ? runs[0]
          : runs.find((candidate) => candidate.runId.startsWith(run));
      if (!listing) {
        throw new ConsultChimpsError(
          "CLI_LOG_NOT_FOUND",
          run === "latest"
            ? `No run records yet in ${directory}`
            : `No run record starts with "${run}" in ${directory}`,
        );
      }
      output.data(
        { recordPath: listing.recordPath, summary: listing.summary },
        formatShow(listing),
      );
    });

  logs
    .command("path")
    .description("print the folder that holds run records")
    .action(() => {
      const directory = resolveLogDirectory();
      output.data({ directory }, `${withoutTerminalControls(directory)}\n`);
    });
}
