import { mkdir, stat } from "node:fs/promises";
import path from "node:path";

import { ConsultChimpsError } from "@consultchimps/core";
import fg from "fast-glob";

export interface DiscoverFilesOptions {
  cwd?: string | undefined;
  extensions?: string[] | undefined;
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

export async function discoverFiles(
  inputs: string[],
  options: DiscoverFilesOptions = {},
): Promise<string[]> {
  if (inputs.length === 0) {
    throw new ConsultChimpsError(
      "FILES_NO_INPUTS",
      "At least one input path or pattern is required.",
    );
  }

  const cwd = path.resolve(options.cwd ?? process.cwd());
  const extensions = normalizeExtensions(options.extensions);
  const discovered = new Set<string>();

  for (const input of inputs) {
    const absoluteInput = path.resolve(cwd, input);

    try {
      const inputStat = await stat(absoluteInput);
      if (inputStat.isFile()) {
        discovered.add(absoluteInput);
        continue;
      }

      if (inputStat.isDirectory()) {
        const matches = await fg("**/*", {
          absolute: true,
          cwd: absoluteInput,
          onlyFiles: true,
        });
        matches.forEach((match) => discovered.add(path.resolve(match)));
        continue;
      }
    } catch {
      const matches = await fg(normalizeGlobPattern(input), {
        absolute: true,
        cwd,
        onlyFiles: true,
      });
      matches.forEach((match) => discovered.add(path.resolve(match)));
    }
  }

  const files = [...discovered]
    .filter(
      (filePath) =>
        !extensions || extensions.has(path.extname(filePath).toLowerCase()),
    )
    .sort((left, right) => left.localeCompare(right));

  if (files.length === 0) {
    throw new ConsultChimpsError(
      "FILES_NOT_FOUND",
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
    "FILES_OUTPUT_EXISTS",
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
      "FILES_INPUT_OVERWRITE",
      `Refusing to overwrite an input file: ${resolvedOutput}`,
      { details: { outputPath: resolvedOutput } },
    );
  }
}
