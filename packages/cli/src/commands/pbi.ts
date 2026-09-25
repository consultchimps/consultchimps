import { readFile, rm, rmdir, stat, writeFile } from "node:fs/promises";
import { totalmem } from "node:os";
import path from "node:path";

import {
  ConsultChimpsError,
  isConsultChimpsError,
  type ByteArtifact,
  type OperationResult,
} from "@consultchimps/core";
import {
  discoverFiles,
  ensureDirectory,
  pathExists,
  planFilePublication,
  publishStagedFile,
  type FilePublicationPlan,
} from "@consultchimps/files";
import { exportPbiTables } from "@consultchimps/pbi";
import { InvalidArgumentError, Option, type Command } from "commander";

import { formatPowerBiExport, readExportManifest } from "../pbi-report.js";
import { createCliProgress } from "../progress.js";

/** What the command needs from the program to report its outcome. */
export interface PbiCommandOutput {
  json(): boolean;
  result(value: OperationResult): void;
  prose(value: string): void;
}

interface ExportOptions {
  output?: string;
  includeHidden?: boolean;
  maxMemory: number;
  force?: boolean;
}

/**
 * The two files a run writes, with fixed names inside the output folder. They
 * are fixed rather than derived from the input so a second command, a script,
 * or a colleague knows where to look without repeating the first command's
 * arguments.
 */
const WORKBOOK_FILE = "workbook.xlsx";
const MANIFEST_FILE = "manifest.json";
const WORKBOOK_MEDIA_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const MANIFEST_MEDIA_TYPE = "application/json";

/**
 * The name this command asks the library to build its outputs under, and the two
 * names that produces. Asking for it explicitly is what lets the names be
 * checked: left to the library's own default, they would be whatever a future
 * version defaults to.
 */
const LIBRARY_OUTPUT_NAME = "power-bi-tables.xlsx";
const LIBRARY_WORKBOOK_NAME = "power-bi-tables.xlsx";
const LIBRARY_MANIFEST_NAME = "power-bi-tables.manifest.json";

/**
 * The memory ceiling the command works to.
 *
 * The library default, 768 MiB, is sized for a browser tab that shares its heap
 * with a page. A command-line run has the machine, and the workbook stage is
 * charged at its measured cost, so the default here is 4 GiB: enough for the
 * largest sample model measured, which reserves about 3.84 GB. The floor is
 * 64 MiB because below that the runtimes and the container alone cannot fit, so
 * a smaller ceiling can only produce a refusal.
 *
 * The ceiling is also bounded above, at twice the machine's physical memory. A
 * budget exists so a model too large for the host is refused instead of ending
 * in an out-of-memory, and a ceiling of nine petabytes hands that outcome back:
 * every limit passes and the process dies. Twice physical memory leaves room for
 * a host that pages, and stops short of a number the machine could never reach.
 */
const DEFAULT_MAX_MEMORY = 4 * 1024 * 1024 * 1024;
const MINIMUM_MAX_MEMORY = 64 * 1024 * 1024;
const MAXIMUM_MEMORY_FACTOR = 2;

/**
 * The library's four byte limits, derived from that one ceiling.
 *
 * `peakBytes` is the ceiling itself. The other three are shares of it, because
 * they bound parts of the same run: the file this process holds in memory, what
 * decompression produces from it, and the two outputs it builds before writing
 * them. Leaving them at the library's browser defaults would cap a command-line
 * run at a 64 MiB input whatever the machine has, which is smaller than the
 * Power BI files this command exists for, and it would give a user no way to
 * say so. One ceiling, stated once, governs all four.
 */
function byteLimits(maxMemory: number): {
  inputBytes: number;
  decodedBytes: number;
  outputBytes: number;
  peakBytes: number;
} {
  return {
    inputBytes: Math.floor(maxMemory / 8),
    decodedBytes: Math.floor(maxMemory / 2),
    outputBytes: Math.floor(maxMemory / 4),
    peakBytes: maxMemory,
  };
}

const SIZE_SUFFIXES: Readonly<Record<string, number>> = {
  k: 1024,
  m: 1024 * 1024,
  g: 1024 * 1024 * 1024,
};

const ACCEPTED_SIZES =
  "Give a whole number of bytes, or a number with a k, m, or g suffix, such as 512m or 4g, and at least 64m.";

/** The most this machine will accept, and the sentence that explains it. */
export function maximumMemorySize(): number {
  return MAXIMUM_MEMORY_FACTOR * totalmem();
}

function tooLargeMessage(): string {
  return `That ceiling is above twice the machine's physical memory. The ceiling cannot make the machine larger: a budget above what the host has hands back the out-of-memory it exists to prevent. Give at most ${Math.floor(maximumMemorySize() / (1024 * 1024))}m on this machine.`;
}

/**
 * Read `--max-memory`. Every declared size is bounded before it sizes anything:
 * the text must be digits with at most one suffix, the product must still be a
 * safe integer, it must reach the floor, and it must stay under twice the
 * machine's physical memory. A value that fails any of those is a usage error,
 * which is reported before the file is opened.
 */
export function parseMemorySize(value: string): number {
  const match = /^(\d{1,19})([kmg])?$/i.exec(value.trim());
  if (match === null) throw new InvalidArgumentError(ACCEPTED_SIZES);
  const scale = SIZE_SUFFIXES[(match[2] ?? "b").toLowerCase()] ?? 1;
  const bytes = Number(match[1]) * scale;
  if (!Number.isSafeInteger(bytes) || bytes < MINIMUM_MAX_MEMORY)
    throw new InvalidArgumentError(ACCEPTED_SIZES);
  if (bytes > maximumMemorySize())
    throw new InvalidArgumentError(tooLargeMessage());
  return bytes;
}

/** `<input directory>/<input name>-tables`, when no folder is given. */
export function defaultOutputDirectory(inputPath: string): string {
  return path.join(
    path.dirname(inputPath),
    `${path.parse(inputPath).name}-tables`,
  );
}

export interface StagedOutput {
  readonly plan: FilePublicationPlan;
  readonly bytes: Uint8Array;
  readonly staging: string;
}

/** The publication step, injectable so a test can fail it at either output. */
export type PublishStagedFile = (options: {
  readonly temporary: string;
  readonly plan: FilePublicationPlan;
}) => Promise<void>;

/**
 * Whether a failed publication left the file on disk anyway.
 *
 * `publishStagedFile` puts the file in place first and removes its staging file
 * second, so a failure to remove it is reported with the output already
 * published. Reading that from the error is the only honest way to know what is
 * on disk: a counter incremented after the call returns says the file is absent
 * in exactly the case where it is present.
 */
function publishedDespiteFailure(error: unknown): boolean {
  return (
    isConsultChimpsError(error) &&
    error.code === "FILES_PUBLICATION_CLEANUP_FAILED" &&
    (error.details as { published?: unknown } | undefined)?.published === true
  );
}

/**
 * Publish both finished files, or leave the destination as it was.
 *
 * Both byte arrays exist before either file is staged, each staging file is
 * written whole and exclusively before anything is published, and the staging
 * files this command still owns are removed on every exit, where a removal
 * failure can never replace the error already propagating.
 *
 * What the outcome says is decided by what is on disk, never by how far the loop
 * got. A publication that fails after putting the file in place says so through
 * `FILES_PUBLICATION_CLEANUP_FAILED`, and that file stays counted as published.
 * So there are exactly three unhappy outcomes: nothing published, which is the
 * original failure; one file published and the other not, which is
 * `CLI_PBI_OUTPUT_INCOMPLETE`; and both published with a staging file left
 * behind, which is `CLI_PBI_OUTPUT_CLEANUP_REQUIRED`. Incompleteness outranks
 * litter, because it is the state a reader can act on wrongly.
 */
export async function publishOutputs(
  outputs: readonly StagedOutput[],
  options: { readonly force: boolean },
  publish: PublishStagedFile = publishStagedFile,
): Promise<void> {
  const published: string[] = [];
  const abandonedStaging: string[] = [];
  // Only the staging files this command actually created. A file already at a
  // staging path is somebody else's: the exclusive write below refuses to
  // overwrite it, and it must not be deleted on the way out either.
  const ours = new Set<string>();
  let failure: unknown;
  try {
    for (const output of outputs) {
      await writeFile(output.staging, output.bytes, { flag: "wx" });
      ours.add(output.staging);
    }
    for (const output of outputs) {
      try {
        await publish({ temporary: output.staging, plan: output.plan });
        published.push(output.plan.output);
      } catch (error) {
        if (!publishedDespiteFailure(error)) {
          failure = error;
          break;
        }
        // The file is in place; only its staging file survived, and the files
        // package has already tried to remove it, so this command does not.
        published.push(output.plan.output);
        abandonedStaging.push(output.staging);
        ours.delete(output.staging);
      }
    }
  } finally {
    for (const output of outputs)
      if (ours.has(output.staging))
        // Litter this command owns. Neither removal may mask a failure above.
        await rm(output.staging, { force: true }).catch(() => {});
  }

  if (failure !== undefined) {
    if (published.length === 0) throw failure;
    const missing = outputs
      .filter((output) => !published.includes(output.plan.output))
      .map((output) => output.plan.output);
    throw new ConsultChimpsError(
      "CLI_PBI_OUTPUT_INCOMPLETE",
      `The export wrote ${published.join(" and ")} but could not write ${missing.join(" and ")}, so the pair on disk is incomplete.${
        options.force
          ? " Because replacement was allowed, what stands beside the new file may be the previous run's, describing different tables."
          : ""
      } Remove the file that was written, or rerun the command with --force once the destination is writable.${
        abandonedStaging.length === 0
          ? ""
          : ` The staging file ${abandonedStaging.join(" and ")} also remains and can be deleted.`
      }`,
      {
        details: {
          published,
          missing,
          ...(abandonedStaging.length === 0
            ? {}
            : { stagingFiles: abandonedStaging }),
        },
      },
    );
  }

  if (abandonedStaging.length > 0)
    throw new ConsultChimpsError(
      "CLI_PBI_OUTPUT_CLEANUP_REQUIRED",
      `Both files were written, but the staging file ${abandonedStaging.join(" and ")} could not be removed afterwards. The export itself is complete. Close whatever is holding that file and delete it.`,
      {
        details: {
          published,
          stagingFiles: abandonedStaging,
          operationCompleted: true,
        },
      },
    );
}

/**
 * The outputs this command knows how to write, checked before any of them is.
 *
 * Names as well as media types: the library documents the workbook first and the
 * manifest second, and this command writes them under names that say which is
 * which, so a swap or a rename has to stop the run rather than put JSON in a
 * file called workbook.xlsx.
 */
export function verifyExportOutputs(
  outputs: readonly ByteArtifact[],
  names: readonly string[] = [LIBRARY_WORKBOOK_NAME, LIBRARY_MANIFEST_NAME],
  mediaTypes: readonly string[] = [WORKBOOK_MEDIA_TYPE, MANIFEST_MEDIA_TYPE],
): void {
  const matches =
    outputs.length === names.length &&
    outputs.every(
      (artifact, index) =>
        artifact.name === names[index] &&
        artifact.mediaType === mediaTypes[index],
    );
  if (!matches)
    throw new ConsultChimpsError(
      "CLI_PBI_UNEXPECTED_OUTPUTS",
      "The export returned outputs this command does not recognize, so nothing was written. Update the command-line interface and the Power BI package together.",
    );
}

export function registerPbiCommands(
  program: Command,
  output: PbiCommandOutput,
): void {
  const pbi = program
    .command("pbi")
    .description(
      "export the tables stored inside a Power BI file, without opening Power BI Desktop",
    )
    .addHelpText(
      "after",
      `
Example:
  consultchimps pbi export sales.pbix -o sales-tables

Safety:
  Your original Power BI file is not changed. ConsultChimps writes new files and
  refuses to replace existing outputs unless you use --force.

Run consultchimps pbi help <command> for all command options.
`,
    );

  pbi
    .command("export")
    .description(
      "write every exportable table in a .pbix model to one Excel workbook, with a manifest of what was left out",
    )
    .argument("<input>", "the source .pbix file")
    .option(
      "-o, --output <directory>",
      "folder for the new workbook.xlsx and manifest.json; defaults to a folder beside the input",
    )
    .option(
      "--include-hidden",
      "export the tables the model marks hidden as well as the visible ones",
    )
    .addOption(
      new Option(
        "--max-memory <size>",
        "most memory the whole export may use, such as 512m or 4g",
      )
        .argParser(parseMemorySize)
        // The byte count is what the library receives; the size is what the
        // reader recognizes, so the help shows the size.
        .default(DEFAULT_MAX_MEMORY, "4g"),
    )
    .option(
      "-f, --force",
      "replace an existing workbook.xlsx or manifest.json in the output folder; use with care",
    )
    .addHelpText(
      "after",
      `
Examples:
  consultchimps pbi export sales.pbix
  consultchimps pbi export sales.pbix -o sales-tables --include-hidden

What happens:
  ConsultChimps reads the model inside the file, writes one worksheet for each
  table, splitting a table too tall for one worksheet across numbered
  worksheets, and writes a manifest naming every table and column it left out
  and why. The workbook carries the model's stored values, not Power BI's
  formatting. Files on a live connection or DirectQuery, and templates, hold no
  imported rows and are refused.
`,
    )
    .action(async (input: string, options: ExportOptions) => {
      const inputPaths = await discoverFiles([input], {
        extensions: [".pbix"],
      });
      const inputPath = inputPaths[0];
      if (inputPath === undefined) {
        // discoverFiles raises FILES_NOT_FOUND itself when nothing matches, so
        // this is the belt to that braces. It carries the same published code,
        // because an empty result means the same thing to a reader and to the
        // --json envelope, which has no shape for a failure without one.
        throw new ConsultChimpsError(
          "FILES_NOT_FOUND",
          "No Power BI file matched the input. Check the path and that the file has a .pbix extension.",
        );
      }
      // One run writes one workbook under a fixed name, so a pattern that
      // matches several files is a request this command cannot carry out. It
      // says so rather than exporting whichever file sorted first.
      if (inputPaths.length > 1) {
        throw new ConsultChimpsError(
          "CLI_PBI_ONE_INPUT",
          `The input matched ${inputPaths.length} Power BI files. This command exports one file at a time, because both outputs have fixed names inside the output folder. Name a single .pbix file, or run the command once for each.`,
        );
      }
      const outputDirectory = path.resolve(
        options.output ?? defaultOutputDirectory(inputPath),
      );
      // Both destinations are checked before anything is decoded, so a run that
      // would refuse to replace an existing file refuses in a second rather
      // than after minutes of work.
      const plans: FilePublicationPlan[] = [];
      for (const name of [WORKBOOK_FILE, MANIFEST_FILE])
        plans.push(
          await planFilePublication({
            output: path.join(outputDirectory, name),
            inputs: [inputPath],
            overwrite: options.force,
          }),
        );

      // The file's own size is checked against the limit this run will enforce
      // before it is read, so an oversized input is refused instead of being
      // loaded into memory and refused afterwards.
      const limits = byteLimits(options.maxMemory);
      const size = (await stat(inputPath)).size;
      if (size > limits.inputBytes) {
        throw new ConsultChimpsError(
          "PBI_EXPORT_LIMIT_EXCEEDED",
          `This Power BI file is ${size} bytes, above the ${limits.inputBytes} bytes this run may hold. The memory ceiling is shared out: the input file may use one eighth of it, decompression one half, and the two outputs one quarter. Raise --max-memory to at least ${Math.ceil((size * 8) / (1024 * 1024))}m for this file, but only as far as the machine can actually spare.`,
          {
            details: {
              stage: "container",
              option: "inputBytes",
              share: "one eighth of --max-memory",
              limit: limits.inputBytes,
              required: size,
            },
          },
        );
      }

      const progress = createCliProgress(output.json());
      let outcome;
      try {
        outcome = await exportPbiTables(await readFile(inputPath), {
          ...limits,
          includeHiddenTables: options.includeHidden === true,
          outputName: LIBRARY_OUTPUT_NAME,
          onProgress: progress.report,
        });
      } finally {
        progress.finish();
      }

      verifyExportOutputs(outcome.outputs);
      // The manifest is read back before anything is written. A failure here is
      // a defect in this pair of packages rather than bad input, and a defect
      // must refuse with nothing on disk rather than exit non-zero after two
      // correct files have been published.
      const report = readExportManifest(outcome.outputs[1]!.bytes);

      const directoryExisted = await pathExists(outputDirectory);
      await ensureDirectory(outputDirectory);
      try {
        await publishOutputs(
          outcome.outputs.map((artifact, index) => ({
            plan: plans[index]!,
            bytes: artifact.bytes,
            // The process id keeps this command's staging file apart from a
            // second run's; the exclusive write keeps it off anyone else's file.
            staging: `${plans[index]!.output}.${process.pid}.part`,
          })),
          { force: options.force === true },
        );
      } catch (error) {
        // A folder this run created and then wrote nothing into is this
        // command's litter. One that already existed, or that holds a published
        // file, is left exactly as it is.
        if (
          !directoryExisted &&
          !(isConsultChimpsError(error) && "published" in (error.details ?? {}))
        )
          await rmdir(outputDirectory).catch(() => {});
        throw error;
      }

      // The library names its artifacts for a host that has no filesystem; the
      // result a command reports has to name the files it actually wrote.
      const result: OperationResult = {
        ...outcome.result,
        artifacts: outcome.outputs.map((artifact, index) => ({
          kind: "file" as const,
          path: plans[index]!.output,
          mediaType: artifact.mediaType!,
        })),
      };
      if (!output.json()) output.prose(`${formatPowerBiExport(report)}\n`);
      output.result(result);
    });
}
