import { copyFile, mkdtemp, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { ConsultChimpsError, isConsultChimpsError } from "@consultchimps/core";
import { planFilePublication } from "@consultchimps/files";
import { afterEach, describe, expect, it } from "vitest";

import {
  publishOutputs,
  verifyExportOutputs,
  type PublishStagedFile,
  type StagedOutput,
} from "../src/commands/pbi.js";

/**
 * What the command says, and what is on disk, when publishing goes wrong.
 *
 * The two outputs are published one after the other, so a failure can leave one
 * file in place and not the other, and a publication can also succeed and still
 * fail, when the file is in place but its staging file cannot be removed. Each
 * case below injects one of those at one step through the publish seam and
 * asserts both halves of the outcome: the code the user is given, and the exact
 * set of files the folder holds afterwards.
 */

const temporaryDirectories: string[] = [];

afterEach(async () => {
  await Promise.all(
    temporaryDirectories
      .splice(0)
      .map((directory) => rm(directory, { force: true, recursive: true })),
  );
});

async function workspace(): Promise<string> {
  const directory = await mkdtemp(path.join(tmpdir(), "consultchimps-pub-"));
  temporaryDirectories.push(directory);
  return directory;
}

const workbookBytes = new TextEncoder().encode("workbook");
const manifestBytes = new TextEncoder().encode("{}");

async function stagedOutputs(directory: string): Promise<StagedOutput[]> {
  const outputs: StagedOutput[] = [];
  for (const [name, bytes] of [
    ["workbook.xlsx", workbookBytes],
    ["manifest.json", manifestBytes],
  ] as const)
    outputs.push({
      plan: await planFilePublication({
        output: path.join(directory, name),
        inputs: [],
        overwrite: true,
      }),
      bytes,
      staging: path.join(directory, `${name}.test.part`),
    });
  return outputs;
}

/** A publication that puts the file in place, exactly as the real one does. */
const publishing: PublishStagedFile = async ({ temporary, plan }) => {
  await copyFile(temporary, plan.output);
  await rm(temporary, { force: true });
};

/** The real failure the files package raises when the file is already in place. */
function cleanupFailure(output: string, temporary: string): ConsultChimpsError {
  return new ConsultChimpsError(
    "FILES_PUBLICATION_CLEANUP_FAILED",
    "The output was published, but its private staging file could not be removed.",
    { details: { published: true, output, temporary } },
  );
}

/** Publish normally except at one step, where the given failure happens. */
function failingAt(
  step: number,
  failure: (output: StagedOutput) => Promise<never>,
): PublishStagedFile {
  let seen = 0;
  return async (options) => {
    const current = seen++;
    if (current !== step) return publishing(options);
    return failure({
      plan: options.plan,
      staging: options.temporary,
      bytes: new Uint8Array(),
    });
  };
}

async function refusal(run: () => Promise<unknown>): Promise<unknown> {
  try {
    await run();
  } catch (error) {
    return error;
  }
  throw new Error("the publication was expected to fail");
}

async function contents(directory: string): Promise<string[]> {
  return (await readdir(directory)).sort();
}

describe("publishing both outputs", () => {
  it("writes both files and removes both staging files", async () => {
    const directory = await workspace();
    await publishOutputs(await stagedOutputs(directory), { force: false });
    expect(await contents(directory)).toEqual([
      "manifest.json",
      "workbook.xlsx",
    ]);
  });

  it("reports the original failure and leaves nothing when the workbook fails", async () => {
    const directory = await workspace();
    const outputs = await stagedOutputs(directory);
    const error = await refusal(() =>
      publishOutputs(
        outputs,
        { force: false },
        failingAt(0, () => {
          throw new ConsultChimpsError("FILES_DESTINATION_CHANGED", "changed");
        }),
      ),
    );
    expect(isConsultChimpsError(error) && error.code).toBe(
      "FILES_DESTINATION_CHANGED",
    );
    // Nothing published, and this command's own staging files are gone.
    expect(await contents(directory)).toEqual([]);
  });

  it("reports a cleanup failure after the workbook, and finishes the pair", async () => {
    const directory = await workspace();
    const outputs = await stagedOutputs(directory);
    const error = await refusal(() =>
      publishOutputs(
        outputs,
        { force: false },
        failingAt(0, async (output) => {
          // The real shape: the file is in place, the staging file is not.
          await copyFile(output.staging, output.plan.output);
          throw cleanupFailure(output.plan.output, output.staging);
        }),
      ),
    );
    expect(isConsultChimpsError(error) && error.code).toBe(
      "CLI_PBI_OUTPUT_CLEANUP_REQUIRED",
    );
    expect(isConsultChimpsError(error) && error.message).toContain(
      "workbook.xlsx.test.part",
    );
    // Both outputs written, and the staging file the files package could not
    // remove is still there, named in the refusal.
    expect(await contents(directory)).toEqual([
      "manifest.json",
      "workbook.xlsx",
      "workbook.xlsx.test.part",
    ]);
  });

  it("reports an incomplete pair when the manifest fails", async () => {
    const directory = await workspace();
    const outputs = await stagedOutputs(directory);
    const error = await refusal(() =>
      publishOutputs(
        outputs,
        { force: false },
        failingAt(1, () => {
          throw new ConsultChimpsError("FILES_DESTINATION_CHANGED", "changed");
        }),
      ),
    );
    expect(isConsultChimpsError(error) && error.code).toBe(
      "CLI_PBI_OUTPUT_INCOMPLETE",
    );
    const message = isConsultChimpsError(error) ? error.message : "";
    expect(message).toContain("workbook.xlsx");
    expect(message).toContain("manifest.json");
    // The workbook stands alone, and no staging file survives it.
    expect(await contents(directory)).toEqual(["workbook.xlsx"]);
  });

  it("reports a cleanup failure after the manifest", async () => {
    const directory = await workspace();
    const outputs = await stagedOutputs(directory);
    const error = await refusal(() =>
      publishOutputs(
        outputs,
        { force: false },
        failingAt(1, async (output) => {
          await copyFile(output.staging, output.plan.output);
          throw cleanupFailure(output.plan.output, output.staging);
        }),
      ),
    );
    expect(isConsultChimpsError(error) && error.code).toBe(
      "CLI_PBI_OUTPUT_CLEANUP_REQUIRED",
    );
    expect(await contents(directory)).toEqual([
      "manifest.json",
      "manifest.json.test.part",
      "workbook.xlsx",
    ]);
  });

  it("says a stale file may remain when replacement was allowed", async () => {
    const directory = await workspace();
    // A previous export, which --force is about to replace.
    await writeFile(path.join(directory, "workbook.xlsx"), "old workbook");
    await writeFile(path.join(directory, "manifest.json"), "old manifest");
    const outputs = await stagedOutputs(directory);
    const error = await refusal(() =>
      publishOutputs(
        outputs,
        { force: true },
        failingAt(1, () => {
          throw new ConsultChimpsError("FILES_DESTINATION_CHANGED", "changed");
        }),
      ),
    );
    expect(isConsultChimpsError(error) && error.code).toBe(
      "CLI_PBI_OUTPUT_INCOMPLETE",
    );
    // The point of the --force wording: the manifest beside the new workbook is
    // the previous run's, and the message has to say so.
    expect(isConsultChimpsError(error) && error.message).toContain(
      "the previous run's",
    );
    expect(await contents(directory)).toEqual([
      "manifest.json",
      "workbook.xlsx",
    ]);
  });

  it("refuses to write over a file already at a staging path", async () => {
    const directory = await workspace();
    const outputs = await stagedOutputs(directory);
    await writeFile(outputs[0]!.staging, "someone else's file");
    const error = await refusal(() =>
      publishOutputs(outputs, { force: false }),
    );
    expect((error as NodeJS.ErrnoException).code).toBe("EEXIST");
    // Their file is untouched, and nothing was published.
    expect(await contents(directory)).toEqual(["workbook.xlsx.test.part"]);
  });
});

describe("the outputs this command knows how to write", () => {
  const workbook = {
    name: "power-bi-tables.xlsx",
    bytes: workbookBytes,
    mediaType:
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  };
  const manifest = {
    name: "power-bi-tables.manifest.json",
    bytes: manifestBytes,
    mediaType: "application/json",
  };

  it("accepts the pair the library documents", () => {
    expect(() => verifyExportOutputs([workbook, manifest])).not.toThrow();
  });

  it.each([
    ["swapped", [manifest, workbook]],
    ["renamed", [{ ...workbook, name: "tables.xlsx" }, manifest]],
    ["one output", [workbook]],
    ["three outputs", [workbook, manifest, manifest]],
    [
      "the right names with the wrong types",
      [
        { ...workbook, mediaType: "application/json" },
        { ...manifest, mediaType: "text/plain" },
      ],
    ],
  ])("refuses %s", (_label, outputs) => {
    let code: string | undefined;
    try {
      verifyExportOutputs(outputs);
    } catch (error) {
      code = isConsultChimpsError(error) ? error.code : undefined;
    }
    expect(code).toBe("CLI_PBI_UNEXPECTED_OUTPUTS");
  });
});
