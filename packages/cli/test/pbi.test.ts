import { execFile } from "node:child_process";
import { createHash } from "node:crypto";
import {
  copyFile,
  mkdtemp,
  mkdir,
  readFile,
  readdir,
  rm,
  writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

import { afterEach, describe, expect, it } from "vitest";

import { maximumMemorySize, parseMemorySize } from "../src/commands/pbi.js";

/**
 * `consultchimps pbi export`, run as the built binary the way a user runs it.
 *
 * The committed Power BI fixture is the input for every case here: it is the
 * one sample whose cells the repository is willing to hold, and it carries a
 * hidden table, so one run exercises the export, an exclusion, and the warning
 * that explains it.
 */

const execFileAsync = promisify(execFile);
const cliPath = fileURLToPath(new URL("../dist/index.js", import.meta.url));
const pbixFixture = fileURLToPath(
  new URL("../../pbi/fixtures/a-2018-fuzzy.pbix", import.meta.url),
);
const temporaryDirectories: string[] = [];

interface CliResult {
  exitCode: number;
  stderr: string;
  stdout: string;
}

async function runCli(
  args: string[],
  expectedExitCode = 0,
): Promise<CliResult> {
  let outcome: CliResult;
  try {
    const result = await execFileAsync(process.execPath, [cliPath, ...args], {
      encoding: "utf8",
      maxBuffer: 32 * 1024 * 1024,
    });
    outcome = { exitCode: 0, stderr: result.stderr, stdout: result.stdout };
  } catch (error) {
    const processError = error as Error & {
      code?: number;
      stderr?: string;
      stdout?: string;
    };
    outcome = {
      exitCode: typeof processError.code === "number" ? processError.code : 1,
      stderr: processError.stderr ?? "",
      stdout: processError.stdout ?? "",
    };
  }
  expect(
    outcome.exitCode,
    outcome.stderr || `Command exited with ${outcome.exitCode}`,
  ).toBe(expectedExitCode);
  return outcome;
}

async function createWorkspace(): Promise<{
  directory: string;
  input: string;
}> {
  const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-pbi-"));
  temporaryDirectories.push(directory);
  await mkdir(path.join(directory, "inputs"));
  const input = path.join(directory, "inputs", "model.pbix");
  await copyFile(pbixFixture, input);
  return { directory, input };
}

async function digest(filePath: string): Promise<string> {
  return createHash("sha256")
    .update(await readFile(filePath))
    .digest("hex");
}

interface JsonSuccess {
  ok: true;
  result: {
    operation: string;
    artifacts: { kind: string; path: string; mediaType: string }[];
    metrics: Record<string, number>;
    warnings: string[];
  };
}

interface JsonFailure {
  ok: false;
  error: { code: string | null; message: string };
}

/** One JSON object on one line of stdout is itself part of the contract. */
function parseSingleJsonLine(stdout: string): unknown {
  expect(stdout.endsWith("\n")).toBe(true);
  const body = stdout.slice(0, -1);
  expect(body).not.toContain("\n");
  return JSON.parse(body);
}

afterEach(async () => {
  await Promise.all(
    temporaryDirectories
      .splice(0)
      .map((directory) => rm(directory, { force: true, recursive: true })),
  );
});

describe("consultchimps pbi export", () => {
  it("writes a workbook and a manifest beside the input, and repeats byte for byte", async () => {
    const first = await createWorkspace();
    const firstFolder = path.join(first.directory, "inputs", "model-tables");
    const result = await runCli(["pbi", "export", first.input]);

    expect(result.stdout).toContain("Tables exported:");
    expect(result.stdout).toContain("SUCCESS");
    // The manifest names the hidden table that was left out; the warning
    // explains it in plain language.
    expect(result.stdout).toContain("Tables left out:");
    expect(result.stdout).toContain("PBI_TABLE_HIDDEN");
    expect(result.stdout).toContain("Power BI export manifest");

    const workbook = path.join(firstFolder, "workbook.xlsx");
    const manifest = path.join(firstFolder, "manifest.json");
    const before = [await digest(workbook), await digest(manifest)];

    // A second run from the same bytes into a different folder: identical
    // outputs are the contract, so the comparison is of both files' digests.
    const second = await createWorkspace();
    await runCli([
      "pbi",
      "export",
      second.input,
      "-o",
      path.join(second.directory, "out"),
    ]);
    const after = [
      await digest(path.join(second.directory, "out", "workbook.xlsx")),
      await digest(path.join(second.directory, "out", "manifest.json")),
    ];
    expect(after).toEqual(before);

    // The manifest is readable JSON with the schema the library documents.
    const document = JSON.parse(await readFile(manifest, "utf8")) as {
      schemaVersion: number;
      tables: { name: string; parts: { sheetName: string }[] }[];
      excludedTables: { reasons: { code: string }[] }[];
    };
    expect(document.schemaVersion).toBe(1);
    expect(document.tables.length).toBeGreaterThan(0);
    expect(
      document.excludedTables.flatMap((table) =>
        table.reasons.map((reason) => reason.code),
      ),
    ).toContain("PBI_TABLE_HIDDEN");

    // Exactly the two files, and no staging file left beside them: the writer
    // publishes both or neither and cleans up after itself either way.
    expect((await readdir(firstFolder)).sort()).toEqual([
      "manifest.json",
      "workbook.xlsx",
    ]);
  });

  it("prints the outcome envelope and nothing else with --json", async () => {
    const workspace = await createWorkspace();
    const folder = path.join(workspace.directory, "out");
    const result = await runCli([
      "--json",
      "pbi",
      "export",
      workspace.input,
      "-o",
      folder,
    ]);
    expect(result.stderr).toBe("");
    const envelope = parseSingleJsonLine(result.stdout) as JsonSuccess;
    expect(envelope.ok).toBe(true);
    expect(envelope.result.operation).toBe("pbi.export");
    expect(envelope.result.artifacts.map((artifact) => artifact.path)).toEqual([
      path.join(folder, "workbook.xlsx"),
      path.join(folder, "manifest.json"),
    ]);
    expect(
      envelope.result.artifacts.map((artifact) => artifact.mediaType),
    ).toEqual([
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "application/json",
    ]);
    expect(Object.keys(envelope.result.metrics)).toEqual([
      "inputFiles",
      "outputFiles",
      "exportedTables",
      "exportedColumns",
      "exportedRows",
      "outputWorksheets",
    ]);
    expect(envelope.result.warnings.length).toBeGreaterThan(0);
  });

  it("exports the hidden tables only when asked", async () => {
    const workspace = await createWorkspace();
    const plain = await runCli([
      "--json",
      "pbi",
      "export",
      workspace.input,
      "-o",
      path.join(workspace.directory, "plain"),
    ]);
    const hidden = await runCli([
      "--json",
      "pbi",
      "export",
      workspace.input,
      "--include-hidden",
      "-o",
      path.join(workspace.directory, "hidden"),
    ]);
    const exported = (result: CliResult): number =>
      (parseSingleJsonLine(result.stdout) as JsonSuccess).result.metrics[
        "exportedTables"
      ]!;
    expect(exported(hidden)).toBeGreaterThan(exported(plain));
  });

  it("refuses to replace either existing output unless forced", async () => {
    const workspace = await createWorkspace();
    const folder = path.join(workspace.directory, "out");
    await runCli(["pbi", "export", workspace.input, "-o", folder]);

    const refused = await runCli(
      ["--json", "pbi", "export", workspace.input, "-o", folder],
      1,
    );
    const failure = parseSingleJsonLine(refused.stdout) as JsonFailure;
    expect(failure.ok).toBe(false);
    expect(failure.error.code).toBe("FILES_OUTPUT_EXISTS");

    // The manifest alone is enough to refuse: both destinations are checked.
    await rm(path.join(folder, "workbook.xlsx"));
    const manifestOnly = await runCli(
      ["--json", "pbi", "export", workspace.input, "-o", folder],
      1,
    );
    expect(
      (parseSingleJsonLine(manifestOnly.stdout) as JsonFailure).error.code,
    ).toBe("FILES_OUTPUT_EXISTS");

    const forced = await runCli([
      "--json",
      "pbi",
      "export",
      workspace.input,
      "-o",
      folder,
      "--force",
    ]);
    expect((parseSingleJsonLine(forced.stdout) as JsonSuccess).ok).toBe(true);
  });

  it("leaves nothing behind when it refuses", async () => {
    const workspace = await createWorkspace();
    const folder = path.join(workspace.directory, "out");
    const damaged = path.join(workspace.directory, "inputs", "damaged.pbix");
    await writeFile(damaged, "this is not a Power BI file");

    await runCli(["pbi", "export", damaged, "-o", folder], 1);
    // No folder, no staging file, no half-written workbook.
    await expect(
      readFile(path.join(folder, "workbook.xlsx")),
    ).rejects.toThrow();
    await expect(
      readFile(path.join(folder, "manifest.json")),
    ).rejects.toThrow();
  });

  it("renders a library refusal through the shared message wording", async () => {
    const workspace = await createWorkspace();
    const damaged = path.join(workspace.directory, "inputs", "damaged.pbix");
    await writeFile(damaged, "this is not a Power BI file");

    const human = await runCli(
      ["pbi", "export", damaged, "-o", path.join(workspace.directory, "out")],
      1,
    );
    expect(human.stderr).toContain(
      "ERROR: ConsultChimps could not finish your task.",
    );
    expect(human.stderr).toContain("What you can do:");
    expect(human.stderr).toContain(
      "Confirm the file is a Power BI .pbix, not a renamed, truncated, or partly downloaded copy.",
    );
    expect(human.stderr).toContain("PBI_INVALID_CONTAINER");
    expect(human.stderr).not.toContain(
      "keep the error reference below when asking for support",
    );

    const machine = await runCli(
      [
        "--json",
        "pbi",
        "export",
        damaged,
        "-o",
        path.join(workspace.directory, "out"),
      ],
      1,
    );
    const failure = parseSingleJsonLine(machine.stdout) as JsonFailure;
    expect(failure.error.code).toBe("PBI_INVALID_CONTAINER");
  });

  it("refuses a pattern that matches several Power BI files", async () => {
    const workspace = await createWorkspace();
    await copyFile(
      workspace.input,
      path.join(workspace.directory, "inputs", "second.pbix"),
    );
    const result = await runCli(
      [
        "--json",
        "pbi",
        "export",
        path.join(workspace.directory, "inputs", "*.pbix"),
        "-o",
        path.join(workspace.directory, "out"),
      ],
      1,
    );
    expect((parseSingleJsonLine(result.stdout) as JsonFailure).error.code).toBe(
      "CLI_PBI_ONE_INPUT",
    );
  });

  it("refuses an input that is not a Power BI file before reading anything", async () => {
    const workspace = await createWorkspace();
    const notPbix = path.join(workspace.directory, "inputs", "notes.txt");
    await writeFile(notPbix, "notes");
    const result = await runCli(
      [
        "--json",
        "pbi",
        "export",
        notPbix,
        "-o",
        path.join(workspace.directory, "out"),
      ],
      1,
    );
    expect((parseSingleJsonLine(result.stdout) as JsonFailure).error.code).toBe(
      "FILES_NOT_FOUND",
    );
  });

  it("says which share of the ceiling an oversized input exceeded", async () => {
    const workspace = await createWorkspace();
    // The floor ceiling gives the input an eighth of 64 MiB, which is 8 MiB, so
    // a file above that is refused from its size before it is read.
    const big = path.join(workspace.directory, "inputs", "big.pbix");
    await writeFile(big, Buffer.alloc(9 * 1024 * 1024));
    const refused = await runCli(
      [
        "--json",
        "pbi",
        "export",
        big,
        "--max-memory",
        "64m",
        "-o",
        path.join(workspace.directory, "out"),
      ],
      1,
    );
    const failure = parseSingleJsonLine(refused.stdout) as JsonFailure;
    expect(failure.error.code).toBe("PBI_EXPORT_LIMIT_EXCEEDED");
    expect(failure.error.message).toContain("one eighth");
    expect(failure.error.message).toContain("--max-memory");
    // Nothing was read, so nothing was written.
    await expect(
      readFile(path.join(workspace.directory, "out", "workbook.xlsx")),
    ).rejects.toThrow();
  });

  it("accepts the documented memory sizes and refuses the rest", async () => {
    const workspace = await createWorkspace();
    for (const size of ["512m", "4g", "134217728"]) {
      const accepted = await runCli([
        "--json",
        "pbi",
        "export",
        workspace.input,
        "--max-memory",
        size,
        "-o",
        path.join(workspace.directory, `out-${size}`),
      ]);
      expect((parseSingleJsonLine(accepted.stdout) as JsonSuccess).ok).toBe(
        true,
      );
    }

    for (const size of ["32m", "0", "big", "4 g", "-1", "1.5g", "4gb"]) {
      const refused = await runCli(
        [
          "--json",
          "pbi",
          "export",
          workspace.input,
          "--max-memory",
          size,
          "-o",
          path.join(workspace.directory, "rejected"),
        ],
        1,
      );
      const failure = parseSingleJsonLine(refused.stdout) as JsonFailure;
      expect(failure.error.code).toBe("CLI_USAGE");
      expect(failure.error.message).toContain("512m or 4g");
    }
  });

  it("passes the ceiling to the library at the floor it accepts", async () => {
    const workspace = await createWorkspace();
    // The committed fixture reserves about 22 MB, which is below the 64 MiB
    // floor this command accepts, so no accepted ceiling can refuse it. What
    // this run proves is that the smallest accepted ceiling is carried into the
    // library and honoured rather than dropped; the refusal itself is covered
    // by the library suite, which can set a ceiling below the floor.
    const folder = path.join(workspace.directory, "out");
    const smallest = await runCli([
      "--json",
      "pbi",
      "export",
      workspace.input,
      "--max-memory",
      "64m",
      "-o",
      folder,
    ]);
    expect((parseSingleJsonLine(smallest.stdout) as JsonSuccess).ok).toBe(true);
  });
});

/**
 * Whatever the command refuses, it refuses without quoting the model. The names
 * below are the committed fixture's own tables and columns, which the manifest
 * of a successful run does carry: a refusal writes no manifest, so none of them
 * may appear in anything a refusal prints, in either mode.
 */
describe("what a refusal is allowed to say", () => {
  const modelNames = [
    "People",
    "Sales",
    "DateTableTemplate",
    "Quantity",
    "metadata.sqlitedb",
    "DataModel",
  ];

  it("names nothing from inside the file, whatever the refusal", async () => {
    const workspace = await createWorkspace();
    const folder = path.join(workspace.directory, "out");
    await runCli(["pbi", "export", workspace.input, "-o", folder]);
    const damaged = path.join(workspace.directory, "inputs", "damaged.pbix");
    await writeFile(damaged, "this is not a Power BI file");
    const notes = path.join(workspace.directory, "inputs", "notes.txt");
    await writeFile(notes, "notes");
    await copyFile(
      workspace.input,
      path.join(workspace.directory, "inputs", "second.pbix"),
    );

    const refusals: string[][] = [
      // The output folder already holds both files.
      ["pbi", "export", workspace.input, "-o", folder],
      // The input is not a readable container.
      ["pbi", "export", damaged, "-o", path.join(workspace.directory, "a")],
      // The input is not a Power BI file at all.
      ["pbi", "export", notes, "-o", path.join(workspace.directory, "b")],
      // The size is not one the command accepts.
      [
        "pbi",
        "export",
        workspace.input,
        "--max-memory",
        "1",
        "-o",
        path.join(workspace.directory, "c"),
      ],
      // The pattern matches more than one file.
      [
        "pbi",
        "export",
        path.join(workspace.directory, "inputs", "*.pbix"),
        "-o",
        path.join(workspace.directory, "d"),
      ],
    ];

    for (const args of refusals)
      for (const mode of [[], ["--json"]]) {
        const result = await runCli([...mode, ...args], 1);
        const printed = `${result.stdout}${result.stderr}`;
        expect(printed.length).toBeGreaterThan(0);
        for (const name of modelNames) expect(printed).not.toContain(name);
      }
  });
});

describe("the --max-memory size parser", () => {
  it.each([
    ["67108864", 67_108_864],
    ["64m", 67_108_864],
    ["512m", 536_870_912],
    ["4g", 4_294_967_296],
    ["4G", 4_294_967_296],
    ["1048576k", 1_073_741_824],
    ["  256m  ", 268_435_456],
  ])("reads %s as %i bytes", (value: string, expected: number) => {
    expect(parseMemorySize(value)).toBe(expected);
  });

  it.each([
    "",
    "0",
    "1",
    "63m",
    "-4g",
    "+4g",
    "4 g",
    "4gb",
    "1.5g",
    "0x40000000",
    "9999999999999999999g",
    "9007199254740993",
    "four gigabytes",
  ])("refuses %s as a usage error", (value: string) => {
    expect(() => parseMemorySize(value)).toThrow();
  });

  it("refuses a ceiling larger than the machine, however safe the integer", () => {
    // The gap the floor alone left open: a safe integer well above the floor is
    // accepted by every other rule, derives limits in the petabytes, and hands
    // back the out-of-memory the budget exists to prevent.
    expect(() => parseMemorySize("9007199254740991")).toThrow(
      /physical memory/,
    );
    expect(() => parseMemorySize(`${maximumMemorySize() + 1}`)).toThrow(
      /physical memory/,
    );
    // The boundary itself is accepted, so the rule is a ceiling and not a gap.
    expect(parseMemorySize(`${maximumMemorySize()}`)).toBe(maximumMemorySize());
  });
});
