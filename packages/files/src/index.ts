import { mkdir, stat } from "node:fs/promises";
import path from "node:path";

import { ConsultChimpsError } from "@consultchimps/core";
import fg from "fast-glob";

export {
  createScratchDirectory,
  openRandomAccessSource,
  type FileSource,
  type ScratchDirectory,
} from "./random-access.js";
export {
  planFilePublication,
  publishStagedFile,
  type FilePublicationPlan,
} from "./publication.js";

/**
 * Stable, published error codes thrown by @consultchimps/files. Values are
 * part of the versioned public API; never change an existing value.
 */
export const FILES_ERRORS = {
  FILES_INPUT_OVERWRITE: "FILES_INPUT_OVERWRITE",
  FILES_NO_INPUTS: "FILES_NO_INPUTS",
  FILES_NOT_FOUND: "FILES_NOT_FOUND",
  FILES_OUTPUT_EXISTS: "FILES_OUTPUT_EXISTS",
} as const;

export type FilesErrorCode = (typeof FILES_ERRORS)[keyof typeof FILES_ERRORS];

export interface DiscoverFilesOptions {
  cwd?: string | undefined;
  extensions?: string[] | undefined;
  /**
   * `"sorted"` (the default) returns every match in alphabetical order of its
   * path. `"given"` keeps the order of `inputs`: a file keeps its place, and a
   * folder or pattern adds its matches, alphabetically, at its place. Either
   * way a file matched twice is returned once, where it first appears.
   */
  order?: "given" | "sorted" | undefined;
}

/**
 * Alphabetical order of two paths. The locale is named because the default one
 * comes from the environment, and a POSIX locale on Linux would order mixed
 * case differently from Windows.
 */
function comparePaths(left: string, right: string): number {
  return left.localeCompare(right, "en");
}

function normalizeExtensions(
  extensions: string[] | undefined,
): Set<string> | undefined {
  if (!extensions || extensions.length === 0) {
    return undefined;
  }

  return new Set(
    extensions.map((extension) => {
      const normalized = extension.toLowerCase();
      return normalized.startsWith(".") ? normalized : `.${normalized}`;
    }),
  );
}

function normalizeGlobPattern(input: string): string {
  return path.sep === "\\" ? input.replaceAll("\\", "/") : input;
}

function filesystemPathKey(filePath: string): string {
  const resolved = path.resolve(filePath);
  return process.platform === "win32" || process.platform === "darwin"
    ? resolved.toLowerCase()
    : resolved;
}

/**
 * Whether two paths name the same file on this platform.
 *
 * Comparing resolved path strings is not enough: Windows and the usual macOS
 * volume fold case, so "Combined.xlsx" and "combined.xlsx" are one file there
 * and two files on Linux. An operation deciding whether two of its
 * destinations collide has to ask the platform's question, not the string's,
 * or it writes one output over another and still reports both.
 */
export function isSameFilesystemPath(
  firstPath: string,
  secondPath: string,
): boolean {
  return filesystemPathKey(firstPath) === filesystemPathKey(secondPath);
}

/**
 * Whether `candidatePath` is `ancestorPath` itself or sits beneath it.
 *
 * Two destinations can collide without being equal: naming a file
 * "report.xlsx" and a second output "report.xlsx/mapping.json" asks for one
 * path to be a file and a directory at once, which no filesystem grants. The
 * check is case-folded like {@link isSameFilesystemPath} for the same reason.
 */
export function isPathWithin(
  candidatePath: string,
  ancestorPath: string,
): boolean {
  const candidate = filesystemPathKey(candidatePath);
  const ancestor = filesystemPathKey(ancestorPath);
  if (candidate === ancestor) {
    return true;
  }
  const relative = path.relative(ancestor, candidate);
  return (
    relative !== "" &&
    !relative.startsWith(`..${path.sep}`) &&
    relative !== ".." &&
    !path.isAbsolute(relative)
  );
}

/** A key naming the file a path reaches, or the path when it cannot be read. */
async function fileIdentity(filePath: string): Promise<string> {
  try {
    const { dev, ino } = await stat(filePath, { bigint: true });
    return ino === 0n ? `path:${filePath}` : `file:${dev}:${ino}`;
  } catch {
    return `path:${filePath}`;
  }
}

export async function discoverFiles(
  inputs: string[],
  options: DiscoverFilesOptions = {},
): Promise<string[]> {
  if (inputs.length === 0) {
    throw new ConsultChimpsError(
      FILES_ERRORS.FILES_NO_INPUTS,
      "At least one input path or pattern is required.",
    );
  }

  const cwd = path.resolve(options.cwd ?? process.cwd());
  const extensions = normalizeExtensions(options.extensions);
  // Keyed by the file itself (device and inode), so a file named in one case
  // and matched by a pattern in another is one file where the volume folds
  // case, and two files that differ only in case stay two where it does not.
  const discovered = new Map<string, string>();

  for (const input of inputs) {
    const absoluteInput = path.resolve(cwd, input);
    let matches: string[] = [];

    try {
      const inputStat = await stat(absoluteInput);
      if (inputStat.isFile()) {
        matches = [absoluteInput];
      } else if (inputStat.isDirectory()) {
        matches = await fg("**/*", {
          absolute: true,
          cwd: absoluteInput,
          onlyFiles: true,
        });
      }
    } catch {
      matches = await fg(normalizeGlobPattern(input), {
        absolute: true,
        cwd,
        onlyFiles: true,
      });
    }
    // Filtered before deduplicating, so an alias with another extension never
    // stands in for the file it reaches.
    for (const match of matches
      .map((candidate) => path.resolve(candidate))
      .filter(
        (candidate) =>
          !extensions || extensions.has(path.extname(candidate).toLowerCase()),
      )
      .sort(comparePaths)) {
      const key = await fileIdentity(match);
      if (!discovered.has(key)) discovered.set(key, match);
    }
  }

  const files = [...discovered.values()];
  if (options.order !== "given") files.sort(comparePaths);

  if (files.length === 0) {
    throw new ConsultChimpsError(
      FILES_ERRORS.FILES_NOT_FOUND,
      "No files matched the supplied inputs.",
      {
        details: { inputs, cwd, extensions: options.extensions },
      },
    );
  }

  return files;
}

export async function pathExists(targetPath: string): Promise<boolean> {
  try {
    await stat(path.resolve(targetPath));
    return true;
  } catch {
    return false;
  }
}

export async function ensureDirectory(directoryPath: string): Promise<string> {
  const absolutePath = path.resolve(directoryPath);
  await mkdir(absolutePath, { recursive: true });
  return absolutePath;
}

export async function ensureParentDirectory(filePath: string): Promise<string> {
  const absolutePath = path.resolve(filePath);
  await ensureDirectory(path.dirname(absolutePath));
  return absolutePath;
}

export async function ensureOutputAvailable(
  outputPath: string,
  options: { overwrite?: boolean | undefined } = {},
): Promise<string> {
  const absolutePath = path.resolve(outputPath);
  if (options.overwrite) {
    return absolutePath;
  }

  try {
    await stat(absolutePath);
  } catch {
    return absolutePath;
  }

  throw new ConsultChimpsError(
    FILES_ERRORS.FILES_OUTPUT_EXISTS,
    `Output already exists: ${absolutePath}`,
    {
      details: { outputPath: absolutePath },
    },
  );
}

export function refuseInputOverwrite(
  outputPath: string,
  inputPaths: string[],
): void {
  const resolvedOutput = path.resolve(outputPath);
  const outputKey = filesystemPathKey(resolvedOutput);
  const inputKeys = new Set(inputPaths.map(filesystemPathKey));

  if (inputKeys.has(outputKey)) {
    throw new ConsultChimpsError(
      FILES_ERRORS.FILES_INPUT_OVERWRITE,
      `Refusing to overwrite an input file: ${resolvedOutput}`,
      { details: { outputPath: resolvedOutput } },
    );
  }
}
