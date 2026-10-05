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
 * Legacy modules that still own ZIP concerns directly.
 *
 * The list is empty: L0 absorbed every one of them. `package-zip.ts` and
 * `package-paths.ts` were folded into `src/package/`; `excel-tables.ts`,
 * `values-only.ts` and `preserve-table-split.ts` reach the archive through
 * `WorkbookPackage`; Phase 1 re-expressed the all-worksheet split on the
 * layers and the Tier-1 utilities on `WorkbookPackage`, which retired the last
 * two entries and the `src/tier1/` exemption below with them.
 *
 * The test asserts this list matches the offending files EXACTLY, in both
 * directions, so a new module cannot quietly join it and a migrated module
 * cannot be left on it. `src/package/` is now the only owner of JSZip.
 */
const JSZIP_ALLOWLIST: readonly string[] = [];

// Every ZIP library the package depends on. The boundary is about owning ZIP
// assembly, not about one library, so a new backend joins this list rather than
// slipping past a check that only knew the old one.
const ZIP_LIBRARIES: readonly string[] = ["jszip", "fflate", "@zip.js/zip.js"];

interface SourceFile {
  /** Path relative to src/, always with forward slashes. */
  readonly relativePath: string;
  readonly imports: readonly string[];
  /** The named bindings of every import clause, concatenated. */
  readonly importedNames: string;
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

/**
 * The `{ ... }` clause of every static import, concatenated. Enough to ask
 * which named bindings a module pulled in, without a parser.
 */
function importedBindings(source: string): string {
  const clauses: string[] = [];
  for (const match of source.matchAll(
    /\bimport\s+(?:type\s+)?\{([^}]*)\}\s*from/gu,
  )) {
    clauses.push(match[1] ?? "");
  }
  return clauses.join(",");
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
    const source = await readFile(absolute, "utf8");
    files.push({
      relativePath,
      imports: moduleSpecifiers(source),
      importedNames: importedBindings(source),
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
    "no file under src/%s/ imports a ZIP library",
    (directory) => {
      const offenders = sourceFiles
        .filter(
          (file) =>
            inDirectory(file, directory) &&
            ZIP_LIBRARIES.some((library) => importsPackage(file, library)),
        )
        .map((file) => file.relativePath);
      expect(offenders).toEqual([]);
    },
  );
});

/**
 * The XML helpers that REWRITE a part. Editing markup is L0/L1 work: an
 * operation that reached for one of these would be the "operations edited raw
 * XML with regexes" problem this architecture exists to end.
 *
 * Read-only helpers are deliberately absent. An operation may decode text it
 * compares against - the inspection decodes a defined name's reference before
 * matching a worksheet name - without rewriting anything.
 */
const XML_MUTATION_HELPERS = [
  "editElements",
  "setAttribute",
  "addAttribute",
  "writeAttribute",
];

describe("boundaries: operations never rewrite XML", () => {
  it("no file under src/operations/ imports an XML mutation helper", () => {
    const offenders = sourceFiles
      .filter((file) => inDirectory(file, "operations"))
      .flatMap((file) =>
        XML_MUTATION_HELPERS.filter((helper) =>
          new RegExp(`\\b${helper}\\b`, "u").test(file.importedNames),
        ).map((helper) => `${file.relativePath}: ${helper}`),
      );
    expect(offenders).toEqual([]);
  });
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
  it("keeps the ZIP library allowlist exact", () => {
    const offenders = sourceFiles
      .filter(
        (file) =>
          ZIP_LIBRARIES.some((library) => importsPackage(file, library)) &&
          // The two ZIP owners ARCHITECTURE.md allows: the package model and
          // the bounded read path.
          !inDirectory(file, "package") &&
          !inDirectory(file, "stream"),
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

const REPOSITORY_DIRECTORY = fileURLToPath(
  new URL("../../../", import.meta.url),
);

/** Every .ts file under each package's src/ and test/, relative to the repository. */
async function repositorySourceFiles(): Promise<string[]> {
  const files: string[] = [];
  for (const workspace of await readdir(
    path.join(REPOSITORY_DIRECTORY, "packages"),
  )) {
    for (const folder of ["src", "test"]) {
      const root = path.join(
        REPOSITORY_DIRECTORY,
        "packages",
        workspace,
        folder,
      );
      const entries = await readdir(root, {
        recursive: true,
        withFileTypes: true,
      }).catch(() => []);
      for (const entry of entries) {
        if (entry.isFile() && /\.[cm]?[jt]sx?$/u.test(entry.name)) {
          files.push(path.join(entry.parentPath, entry.name));
        }
      }
    }
  }
  return files;
}

describe("boundaries: SheetJS stays out of the repository", () => {
  it("no package's src or test imports xlsx", async () => {
    const files = await repositorySourceFiles();
    expect(files.length).toBeGreaterThan(0);
    const offenders: string[] = [];
    for (const file of files) {
      const imports = moduleSpecifiers(await readFile(file, "utf8"));
      if (imports.some((name) => name === "xlsx" || name.startsWith("xlsx/"))) {
        offenders.push(
          path.relative(REPOSITORY_DIRECTORY, file).split(path.sep).join("/"),
        );
      }
    }
    expect(offenders).toEqual([]);
  });

  it("no package.json declares xlsx", async () => {
    const manifests = ["package.json"];
    for (const group of ["packages", "apps"]) {
      for (const workspace of await readdir(
        path.join(REPOSITORY_DIRECTORY, group),
      )) {
        manifests.push(path.join(group, workspace, "package.json"));
      }
    }
    const offenders: string[] = [];
    for (const manifest of manifests) {
      const text = await readFile(
        path.join(REPOSITORY_DIRECTORY, manifest),
        "utf8",
      ).catch(() => undefined);
      if (text === undefined) continue;
      const json = JSON.parse(text) as Record<string, unknown>;
      for (const field of [
        "dependencies",
        "devDependencies",
        "peerDependencies",
        "optionalDependencies",
      ]) {
        const declared = json[field] as Record<string, string> | undefined;
        if (declared !== undefined && "xlsx" in declared) {
          offenders.push(`${manifest.split(path.sep).join("/")}: ${field}`);
        }
      }
    }
    expect(offenders).toEqual([]);
  });
});
