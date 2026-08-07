/**
 * Guardrail 2 of ARCHITECTURE.md: boundary enforcement.
 *
 * "Nothing above a layer reaches below the layer beneath it." The dist
 * import-graph walkers (packages/pdf/test/bytes.test.ts) prove the built bytes
 * entry stays free of node builtins; this test does the same job over SOURCE
 * files, so a fix that lands in the wrong layer fails before it is bundled.
 *
 * Reading the import specifiers with a regular expression is deliberate: this
 * is a test, not the engine, and a real parser would be a dependency bought to
 * check a handful of string literals.
 */
import { readdir, readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { beforeAll, describe, expect, it } from "vitest";

const SOURCE_DIRECTORY = fileURLToPath(new URL("../src/", import.meta.url));

/**
 * Legacy modules that still own ZIP concerns directly. ARCHITECTURE.md's L0
 * section names the first three explicitly ("the previous parallel
 * implementations ... converge here"); the other two are the readers and
 * writers those three are built on.
 *
 * The test asserts this list matches the offending files EXACTLY, so a new
 * module cannot quietly join it.
 */
const JSZIP_ALLOWLIST: readonly string[] = [
  "excel-tables.ts", // migrates in Phase 1 - becomes WorkbookModel.tables()
  "package-zip.ts", // migrates in Phase 1 - becomes the L0 deterministic writer (type-only import today)
  "preserve-table-split.ts", // migrates in Phase 1 - split re-expressed on L0/L1/L2
  "values-only.ts", // migrates in Phase 1 - converges into the L0 package model
  "workbook-column-split.ts", // migrates in Phase 1 - converges into the L0 package model
];

interface SourceFile {
  /** Path relative to src/, always with forward slashes. */
  readonly relativePath: string;
  readonly imports: readonly string[];
}

/**
 * Every module specifier the file imports: static `import`/`export ... from`,
 * bare side-effect imports, and dynamic `import(...)` with a literal argument.
 */
function moduleSpecifiers(source: string): readonly string[] {
  const specifiers: string[] = [];
  const patterns = [
    // Every `import ... from "x"` and `export ... from "x"`, however the
    // clause is wrapped across lines.
    /\bfrom\s*["']([^"']+)["']/gu,
    // Side-effect imports, which carry no `from`.
    /^\s*import\s*["']([^"']+)["']/gmu,
    /\bimport\s*\(\s*["']([^"']+)["']\s*\)/gu,
    /\brequire\s*\(\s*["']([^"']+)["']\s*\)/gu,
  ];
  for (const pattern of patterns) {
    for (const match of source.matchAll(pattern)) {
      specifiers.push(match[1]!);
    }
  }
  return specifiers;
}

async function readSourceFiles(): Promise<readonly SourceFile[]> {
  const entries = await readdir(SOURCE_DIRECTORY, {
    recursive: true,
    withFileTypes: true,
  });
  const files: SourceFile[] = [];
  for (const entry of entries) {
    if (!entry.isFile() || !entry.name.endsWith(".ts")) {
      continue;
    }
    const absolute = path.join(entry.parentPath, entry.name);
    const relativePath = path
      .relative(SOURCE_DIRECTORY, absolute)
      .split(path.sep)
      .join("/");
    files.push({
      relativePath,
      imports: moduleSpecifiers(await readFile(absolute, "utf8")),
    });
  }
  return files;
}

function inDirectory(file: SourceFile, directory: string): boolean {
  return file.relativePath.startsWith(`${directory}/`);
}

function importsPackage(file: SourceFile, packageName: string): boolean {
  return file.imports.some(
    (specifier) =>
      specifier === packageName || specifier.startsWith(`${packageName}/`),
  );
}

let sourceFiles: readonly SourceFile[] = [];

beforeAll(async () => {
  sourceFiles = await readSourceFiles();
});

describe("boundaries: the walker sees the package", () => {
  it("finds source files and their imports", () => {
    expect(sourceFiles.length).toBeGreaterThan(0);
    expect(
      sourceFiles.some((file) => file.imports.length > 0),
      "no imports parsed at all - the specifier patterns have gone stale",
    ).toBe(true);
    // The frozen L0/L1/L2 seams must be where the layer rules expect them.
    for (const seam of [
      "package/types.ts",
      "model/types.ts",
      "region/types.ts",
    ]) {
      expect(sourceFiles.map((file) => file.relativePath)).toContain(seam);
    }
  });
});

describe("boundaries: operations and regions never touch ZIP", () => {
  it.each(["operations", "region"])(
    "no file under src/%s/ imports jszip",
    (directory) => {
      const offenders = sourceFiles
        .filter(
          (file) =>
            inDirectory(file, directory) && importsPackage(file, "jszip"),
        )
        .map((file) => file.relativePath);
      expect(offenders).toEqual([]);
    },
  );
});

describe("boundaries: the model and the regions never touch the filesystem", () => {
  it.each(["model", "region"])(
    "no file under src/%s/ imports node:fs or node:path",
    (directory) => {
      const offenders = sourceFiles
        .filter(
          (file) =>
            inDirectory(file, directory) &&
            (importsPackage(file, "node:fs") ||
              importsPackage(file, "node:path")),
        )
        .map((file) => file.relativePath);
      expect(offenders).toEqual([]);
    },
  );
});

describe("boundaries: src/package/ is the only owner of ZIP concerns", () => {
  it("keeps the jszip allowlist exact", () => {
    const offenders = sourceFiles
      .filter(
        (file) =>
          importsPackage(file, "jszip") && !inDirectory(file, "package"),
      )
      .map((file) => file.relativePath)
      .sort();

    // An exact match in both directions: a new module cannot join the
    // allowlist unnoticed, and a migrated module cannot be left on it.
    expect(offenders).toEqual([...JSZIP_ALLOWLIST].sort());
  });

  it("keeps every allowlisted module at the top of src/", () => {
    // Once a module moves into a layer directory it is no longer legacy, so
    // the allowlist only ever names files that still live at the src/ root.
    for (const entry of JSZIP_ALLOWLIST) {
      expect(entry).not.toContain("/");
    }
  });
});
