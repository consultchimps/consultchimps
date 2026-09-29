import { execFile } from "node:child_process";
import {
  mkdtemp,
  readdir,
  readFile,
  rm,
  stat,
  writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

import { afterEach, describe, expect, it } from "vitest";
import * as XLSX from "xlsx";

import type { TaskLogRecord } from "@consultchimps/core";

const execFileAsync = promisify(execFile);
const cliPath = fileURLToPath(new URL("../dist/index.js", import.meta.url));
const temporaryDirectories: string[] = [];

afterEach(async () => {
  await Promise.all(
    temporaryDirectories
      .splice(0)
      .map((directory) => rm(directory, { force: true, recursive: true })),
  );
});

interface Run {
  exitCode: number;
  stdout: string;
  stderr: string;
}

async function workspace(): Promise<{ directory: string; logs: string }> {
  const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-log-"));
  temporaryDirectories.push(directory);
  return { directory, logs: path.join(directory, "logs") };
}

async function runCli(
  args: string[],
  logs: string,
  env: Record<string, string> = {},
): Promise<Run> {
  try {
    const result = await execFileAsync(process.execPath, [cliPath, ...args], {
      encoding: "utf8",
      env: {
        ...process.env,
        CONSULTCHIMPS_LOG: "on",
        CONSULTCHIMPS_LOG_DIR: logs,
        ...env,
      },
    });
    return { exitCode: 0, stdout: result.stdout, stderr: result.stderr };
  } catch (error) {
    const failed = error as { code?: number; stdout?: string; stderr?: string };
    return {
      exitCode: typeof failed.code === "number" ? failed.code : 1,
      stdout: failed.stdout ?? "",
      stderr: failed.stderr ?? "",
    };
  }
}

async function writeClientWorkbook(
  filePath: string,
  rows: number,
): Promise<void> {
  const data: Array<Array<string | number>> = [["Client", "Amount"]];
  for (let row = 0; row < rows; row += 1) data.push([`Client ${row}`, row]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(data), "Data");
  await writeFile(
    filePath,
    XLSX.write(workbook, { type: "buffer", bookType: "xlsx" }),
  );
}

async function records(logs: string): Promise<TaskLogRecord[][]> {
  const names = (await readdir(logs))
    .filter((name) => name.endsWith(".jsonl"))
    .sort();
  return Promise.all(
    names.map(async (name) =>
      (await readFile(path.join(logs, name), "utf8"))
        .trim()
        .split("\n")
        .map((line) => JSON.parse(line) as TaskLogRecord),
    ),
  );
}

describe("local run records", () => {
  it("records a consolidation's inputs, per-workbook measures, and outcome without names", async () => {
    const { directory, logs } = await workspace();
    await writeClientWorkbook(path.join(directory, "north-client.xlsx"), 3);
    await writeClientWorkbook(path.join(directory, "south-client.xlsx"), 5);

    const run = await runCli(
      [
        "sheets",
        "consolidate",
        path.join(directory, "*.xlsx"),
        "-o",
        path.join(directory, "combined.xlsx"),
      ],
      logs,
    );
    expect(run.exitCode, run.stderr).toBe(0);
    expect(run.stderr).not.toContain("Run log:");

    const [lines] = await records(logs);
    expect(lines).toBeDefined();
    const types = lines!.map((line) => line.type);
    expect(types[0]).toBe("run");
    expect(types.at(-1)).toBe("end");

    const start = lines![0] as Extract<TaskLogRecord, { type: "run" }>;
    expect(start.command).toBe("sheets consolidate");
    expect(start.names).toBe(false);
    expect(start.options.output).toBe(true);

    const inputs = lines!.filter((line) => line.type === "input");
    expect(inputs.map((input) => input.index)).toEqual([1, 2]);
    expect(
      inputs.every((input) => input.name === undefined && input.bytes > 0),
    ).toBe(true);

    const reads = lines!.filter(
      (line): line is Extract<TaskLogRecord, { type: "progress" }> =>
        line.type === "progress" && line.stage === "reading-workbooks",
    );
    expect(reads.map((read) => read.measures?.rows).sort()).toEqual([3, 5]);
    expect(reads.every((read) => read.detail === undefined)).toBe(true);

    const end = lines!.at(-1) as Extract<TaskLogRecord, { type: "end" }>;
    expect(end.outcome).toBe("ok");
    expect(end.metrics?.outputRows).toBe(8);
    expect(end.peakRss).toBeGreaterThan(0);

    // No name, path, or folder from the run appears anywhere in its record.
    const text = JSON.stringify(lines);
    expect(text).not.toContain("north-client");
    expect(text).not.toContain(path.basename(directory));
  });

  it("keeps base names only when asked, and records a failure's code", async () => {
    const { directory, logs } = await workspace();
    await writeClientWorkbook(path.join(directory, "east-client.xlsx"), 2);

    const named = await runCli(
      [
        "--log-names",
        "sheets",
        "inspect",
        path.join(directory, "east-client.xlsx"),
      ],
      logs,
    );
    expect(named.exitCode, named.stderr).toBe(0);

    const failed = await runCli(
      [
        "sheets",
        "consolidate",
        path.join(directory, "missing-client.xlsx"),
        "-o",
        path.join(directory, "out.xlsx"),
      ],
      logs,
    );
    expect(failed.exitCode).toBe(1);
    expect(failed.stderr).toMatch(/Run log: .+\.jsonl/u);

    const [first, second] = await records(logs);
    const input = first!.find((line) => line.type === "input");
    expect(input?.type === "input" && input.name).toBe("east-client.xlsx");
    expect(JSON.stringify(first)).not.toContain(path.basename(directory));

    const end = second!.at(-1) as Extract<TaskLogRecord, { type: "end" }>;
    expect(end.outcome).toBe("error");
    expect(end.error).toEqual({ code: "FILES_NOT_FOUND" });
  });

  it("records an unrecognised extension as blank and keeps option values out", async () => {
    const { directory, logs } = await workspace();
    const oddName = path.join(directory, "Q3.Acme board pack");
    await writeFile(oddName, "not a presentation");

    await runCli(["pptx", "inspect-template", oddName], logs);

    const [lines] = await records(logs);
    const input = lines!.find((line) => line.type === "input");
    expect(input?.type === "input" && input.extension).toBe("");
    expect(JSON.stringify(lines).toLowerCase()).not.toContain("acme");
  });

  it("records an unreadable number as set rather than null", async () => {
    const { directory, logs } = await workspace();
    await writeClientWorkbook(path.join(directory, "client.xlsx"), 2);
    await runCli(
      [
        "sheets",
        "inspect",
        path.join(directory, "client.xlsx"),
        "--samples",
        "nope",
      ],
      logs,
    );
    const [lines] = await records(logs);
    const start = lines![0] as Extract<TaskLogRecord, { type: "run" }>;
    expect(start.options.samples).toBe(true);
  });

  it.skipIf(process.platform === "win32")(
    "keeps records and their folder private to the user",
    async () => {
      const { directory, logs } = await workspace();
      await writeClientWorkbook(path.join(directory, "client.xlsx"), 2);
      await runCli(
        [
          "--cpu-profile",
          "sheets",
          "inspect",
          path.join(directory, "client.xlsx"),
        ],
        logs,
      );
      expect((await stat(logs)).mode & 0o777).toBe(0o700);
      for (const name of await readdir(logs)) {
        expect((await stat(path.join(logs, name))).mode & 0o777).toBe(0o600);
      }
    },
  );

  it("leaves the db commands' own --profile <file> to them", async () => {
    const { directory, logs } = await workspace();
    const run = await runCli(
      [
        "--json",
        "db",
        "import",
        "run",
        path.join(directory, "missing.duckdb"),
        "--input",
        `north=${path.join(directory, "missing.xlsx")}`,
        "--profile",
        path.join(directory, "routing.json"),
      ],
      logs,
    );
    expect(run.exitCode).toBe(1);
    // A global --profile would have taken the flag and left the file name as
    // an extra argument, which Commander reports as a usage error.
    expect(JSON.parse(run.stdout).error.code).not.toBe("CLI_USAGE");
  });

  it("writes nothing when switched off, and leaves --json output alone", async () => {
    const { directory, logs } = await workspace();
    await writeClientWorkbook(path.join(directory, "west-client.xlsx"), 2);
    const input = path.join(directory, "west-client.xlsx");

    await runCli(["--no-log", "sheets", "inspect", input], logs);
    await runCli(["sheets", "inspect", input], logs, {
      CONSULTCHIMPS_LOG: "off",
    });
    await expect(readdir(logs)).rejects.toThrow();

    const json = await runCli(["--json", "sheets", "inspect", input], logs);
    expect(json.exitCode).toBe(0);
    expect(json.stdout.trim().split("\n")).toHaveLength(1);
    expect(JSON.parse(json.stdout).ok).toBe(true);
    const all = await records(logs);
    expect(all).toHaveLength(1);
    // The record does not depend on the output mode.
    const end = all[0]!.at(-1) as Extract<TaskLogRecord, { type: "end" }>;
    expect(end.operation).toBe("sheets.inspect");
    expect(end.metrics).toBeDefined();
  });

  it("lists and summarises runs, and does not record itself", async () => {
    const { directory, logs } = await workspace();
    await writeClientWorkbook(path.join(directory, "client.xlsx"), 2);
    await runCli(
      ["sheets", "inspect", path.join(directory, "client.xlsx")],
      logs,
    );

    const list = await runCli(["logs"], logs);
    expect(list.exitCode, list.stderr).toBe(0);
    expect(list.stdout).toContain("sheets inspect");

    const show = await runCli(["--json", "logs", "show"], logs);
    const summary = JSON.parse(show.stdout).result.summary;
    expect(summary.outcome).toBe("ok");
    expect(summary.run.command).toBe("sheets inspect");

    const where = await runCli(["logs", "path"], logs);
    expect(where.stdout.trim()).toBe(path.resolve(logs));
    expect(await records(logs)).toHaveLength(1);

    const missing = await runCli(["--json", "logs", "show", "19990101"], logs);
    expect(missing.exitCode).toBe(1);
    expect(JSON.parse(missing.stdout).error.code).toBe("CLI_LOG_NOT_FOUND");
  });
});
