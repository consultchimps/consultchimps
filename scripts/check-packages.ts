import { execFileSync } from "node:child_process";
import {
  mkdtempSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

interface PackageMetadata {
  name: string;
  version: string;
  private?: boolean;
}

const workspaceRoot = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
);
const temporaryRoot = mkdtempSync(
  path.join(tmpdir(), "consultchimps-packages-"),
);
const tarballDirectory = path.join(temporaryRoot, "tarballs");
const consumerDirectory = path.join(temporaryRoot, "consumer");
// A second consumer that installs the libraries alone. The command-line package
// depends on the SQLite runtime, because its Power BI export needs it, so the
// full install can no longer show what a library-only install resolves. This
// directory is where that is still true.
const libraryConsumerDirectory = path.join(temporaryRoot, "library-consumer");
const nodeDirectory = path.dirname(process.execPath);
const pnpmCommand = process.platform === "win32" ? process.execPath : "pnpm";
const pnpmArguments =
  process.platform === "win32"
    ? [
        path.join(
          nodeDirectory,
          "node_modules",
          "corepack",
          "dist",
          "corepack.js",
        ),
        "pnpm",
      ]
    : [];
const npmCommand = process.platform === "win32" ? process.execPath : "npm";
const npmArguments =
  process.platform === "win32"
    ? [path.join(nodeDirectory, "node_modules", "npm", "bin", "npm-cli.js")]
    : [];

const packageDirectories = [
  "core",
  "files",
  "tabular",
  "theme",
  "db",
  "messages",
  "pbi",
  "pdf",
  "xlsx",
  "pptx",
  "cli",
] as const;
// The list above is the expected set of publishable packages. It is checked
// against the workspace so that a new publishable package cannot be left out
// of packing, publint, attw, and the consumer install silently, and so that a
// package turned private is removed from the list deliberately.
const publishableDirectories = readdirSync(
  path.join(workspaceRoot, "packages"),
  { withFileTypes: true },
)
  .filter(
    (entry) => entry.isDirectory() && !readPackageMetadata(entry.name).private,
  )
  .map((entry) => entry.name)
  .sort();
const expectedDirectories = [...packageDirectories].sort();
if (
  JSON.stringify(publishableDirectories) !== JSON.stringify(expectedDirectories)
) {
  throw new Error(
    `Publishable workspace packages [${publishableDirectories.join(", ")}] differ from the checked list [${expectedDirectories.join(", ")}]. Update packageDirectories in scripts/check-packages.ts.`,
  );
}
const libraryDirectories = packageDirectories.filter(
  (directory) => directory !== "cli",
);

function readPackageMetadata(directory: string): PackageMetadata {
  const packagePath = path.join(
    workspaceRoot,
    "packages",
    directory,
    "package.json",
  );

  return JSON.parse(readFileSync(packagePath, "utf8")) as PackageMetadata;
}

try {
  mkdirSync(tarballDirectory);
  mkdirSync(consumerDirectory);
  mkdirSync(libraryConsumerDirectory);

  // Pack only the publishable packages listed above. A recursive pack over
  // packages/* also packs private workspace packages, which never publish.
  execFileSync(
    pnpmCommand,
    [
      ...pnpmArguments,
      ...packageDirectories.flatMap((directory) => [
        "--filter",
        `./packages/${directory}`,
      ]),
      "-r",
      "pack",
      "--pack-destination",
      tarballDirectory,
    ],
    {
      cwd: workspaceRoot,
      stdio: "inherit",
    },
  );

  const tarballs = readdirSync(tarballDirectory)
    .filter((filename) => filename.endsWith(".tgz"))
    .map((filename) => path.join(tarballDirectory, filename))
    .sort();

  if (tarballs.length !== packageDirectories.length) {
    throw new Error(
      `Expected ${packageDirectories.length} tarballs, found ${tarballs.length}.`,
    );
  }

  for (const directory of packageDirectories) {
    execFileSync(pnpmCommand, [...pnpmArguments, "exec", "publint"], {
      cwd: path.join(workspaceRoot, "packages", directory),
      stdio: "inherit",
    });
  }

  for (const tarball of tarballs) {
    // The CLI package publishes only a bin entry point, so it has no type
    // resolution surface for arethetypeswrong to analyze.
    if (/^consultchimps-\d/.test(path.basename(tarball))) {
      continue;
    }
    // The Power BI package exports its WebAssembly decoder as an asset, so a
    // bundler can resolve the binary by package path instead of a CDN. That
    // export resolves to neither JavaScript nor type declarations, which is
    // what arethetypeswrong exists to report, so the asset entry point is
    // excluded by name rather than the rule being turned off: every other
    // entry point of this package, and every entry point of every other
    // package, is still analyzed in full.
    const excluded = /^consultchimps-pbi-/.test(path.basename(tarball))
      ? ["--exclude-entrypoints", "./xpress9.wasm"]
      : [];
    execFileSync(
      pnpmCommand,
      [
        ...pnpmArguments,
        "exec",
        "attw",
        tarball,
        "--profile",
        "esm-only",
        ...excluded,
      ],
      {
        cwd: workspaceRoot,
        stdio: "inherit",
      },
    );
  }

  writeFileSync(
    path.join(consumerDirectory, "package.json"),
    `${JSON.stringify(
      {
        name: "consultchimps-package-smoke-test",
        private: true,
        type: "module",
      },
      null,
      2,
    )}\n`,
  );

  execFileSync(
    npmCommand,
    [
      ...npmArguments,
      "install",
      "--ignore-scripts",
      "--no-audit",
      "--no-fund",
      ...tarballs,
    ],
    {
      cwd: consumerDirectory,
      stdio: "inherit",
    },
  );

  const cliMetadata = readPackageMetadata("cli");
  execFileSync(
    npmCommand,
    [...npmArguments, "rebuild", "better-sqlite3", "--no-audit", "--no-fund"],
    {
      cwd: consumerDirectory,
      stdio: "inherit",
    },
  );
  const cliExecutable = path.join(
    consumerDirectory,
    "node_modules",
    ...(process.platform === "win32"
      ? ["consultchimps", "dist", "index.js"]
      : [".bin", "consultchimps"]),
  );
  const installedVersion = execFileSync(
    process.platform === "win32" ? process.execPath : cliExecutable,
    [...(process.platform === "win32" ? [cliExecutable] : []), "--version"],
    {
      cwd: consumerDirectory,
      encoding: "utf8",
    },
  ).trim();

  if (installedVersion !== cliMetadata.version) {
    throw new Error(
      `CLI reported ${installedVersion}; expected ${cliMetadata.version}.`,
    );
  }

  for (const format of ["sqlite", "duckdb"]) {
    const output = path.join(consumerDirectory, `database-smoke.${format}`);
    const command =
      process.platform === "win32" ? process.execPath : cliExecutable;
    const prefix = process.platform === "win32" ? [cliExecutable] : [];
    execFileSync(command, [...prefix, "--json", "db", "create", "-o", output], {
      cwd: consumerDirectory,
      stdio: "pipe",
    });
    const inspection: unknown = JSON.parse(
      execFileSync(command, [...prefix, "--json", "db", "inspect", output], {
        cwd: consumerDirectory,
        encoding: "utf8",
      }),
    );
    if (
      typeof inspection !== "object" ||
      inspection === null ||
      !("ok" in inspection) ||
      inspection.ok !== true
    ) {
      throw new Error(
        `The installed CLI could not reopen its ${format} database.`,
      );
    }
  }

  const packageNames = libraryDirectories.map(
    (directory) => readPackageMetadata(directory).name,
  );
  execFileSync(
    process.execPath,
    [
      "--input-type=module",
      "--eval",
      `await Promise.all(${JSON.stringify(packageNames)}.map((name) => import(name)));`,
    ],
    {
      cwd: consumerDirectory,
      stdio: "inherit",
    },
  );

  // The tree a real user gets, where the command-line package's own dependency
  // on the SQLite runtime sits beside the optional peer. The subpath still has
  // to import there, which is a different question from the library-only tree
  // below, where the peer is genuinely absent.
  execFileSync(
    process.execPath,
    [
      "--input-type=module",
      "--eval",
      `await import("@consultchimps/db/sqlite-read");`,
    ],
    { cwd: consumerDirectory, stdio: "inherit" },
  );

  // A library-only install: the tarballs of every package except the
  // command-line one, so the checks below see exactly what an application that
  // depends on a library alone resolves, including that the optional SQLite
  // runtime is genuinely absent until something asks for it.
  writeFileSync(
    path.join(libraryConsumerDirectory, "package.json"),
    `${JSON.stringify(
      {
        name: "consultchimps-library-smoke-test",
        private: true,
        type: "module",
      },
      null,
      2,
    )}\n`,
  );
  const libraryTarballs = tarballs.filter(
    (tarball) => !/^consultchimps-\d/.test(path.basename(tarball)),
  );
  execFileSync(
    npmCommand,
    [
      ...npmArguments,
      "install",
      "--ignore-scripts",
      "--no-audit",
      "--no-fund",
      ...libraryTarballs,
    ],
    { cwd: libraryConsumerDirectory, stdio: "inherit" },
  );

  // The Power BI reader's synchronous entry, run on the committed fixture from
  // outside the workspace: it proves the tarball carries a usable build, that
  // the fixture reads through an installed copy rather than the source tree, and
  // that the entry needs neither the optional SQLite peer nor a WebAssembly
  // runtime to return the model part.
  const pbiFixture = path.join(
    workspaceRoot,
    "packages",
    "pbi",
    "fixtures",
    "a-2018-fuzzy.pbix",
  );
  execFileSync(
    process.execPath,
    [
      "--input-type=module",
      "--eval",
      `
    import assert from "node:assert/strict";
    import { readFileSync } from "node:fs";
    import { readPbiModelPart } from "@consultchimps/pbi";
    const container = new Uint8Array(readFileSync(${JSON.stringify(pbiFixture)}));
    const model = readPbiModelPart(container);
    assert.ok(model instanceof Uint8Array);
    assert.ok(model.byteLength > 0);
    assert.notEqual(model.buffer, container.buffer);
  `,
    ],
    { cwd: libraryConsumerDirectory, stdio: "inherit" },
  );

  // The asset export, from an installed copy. A published package whose wasm
  // file is missing from "files", or whose export points somewhere else, passes
  // every other check in this script and fails in a browser bundler, so the
  // subpath is resolved and the file behind it measured.
  execFileSync(
    process.execPath,
    [
      "--input-type=module",
      "--eval",
      `
    import assert from "node:assert/strict";
    import { statSync } from "node:fs";
    import { fileURLToPath } from "node:url";
    const resolved = import.meta.resolve("@consultchimps/pbi/xpress9.wasm");
    assert.ok(resolved.startsWith("file:"), "the wasm export must resolve to an installed file");
    assert.ok(statSync(fileURLToPath(resolved)).size > 0, "the installed xpress9.wasm is empty");
  `,
    ],
    { cwd: libraryConsumerDirectory, stdio: "inherit" },
  );

  // This subpath must import without the optional WASM peer installed. Loading
  // the reader then uses that peer explicitly, from a non-workspace directory.
  execFileSync(
    process.execPath,
    [
      "--input-type=module",
      "--eval",
      `
    import assert from "node:assert/strict";
    await assert.rejects(import("@sqlite.org/sqlite-wasm"), { code: "ERR_MODULE_NOT_FOUND" });
    await import("@consultchimps/db/sqlite-read");
  `,
    ],
    { cwd: libraryConsumerDirectory, stdio: "inherit" },
  );
  const sqlitePeer = (
    JSON.parse(
      readFileSync(
        path.join(workspaceRoot, "packages/db/package.json"),
        "utf8",
      ),
    ) as { peerDependencies: Record<string, string> }
  ).peerDependencies["@sqlite.org/sqlite-wasm"]!;
  execFileSync(
    npmCommand,
    [
      ...npmArguments,
      "install",
      "--ignore-scripts",
      "--no-audit",
      "--no-fund",
      `@sqlite.org/sqlite-wasm@${sqlitePeer}`,
    ],
    { cwd: libraryConsumerDirectory, stdio: "inherit" },
  );
  execFileSync(
    process.execPath,
    [
      "--input-type=module",
      "--eval",
      `
    import assert from "node:assert/strict";
    import initialize from "@sqlite.org/sqlite-wasm";
    import { openReadOnlySqlite } from "@consultchimps/db/sqlite-read";
    const sqlite = await initialize();
    const source = new sqlite.oo1.DB(":memory:");
    source.exec("CREATE TABLE example (id INTEGER); INSERT INTO example VALUES(9223372036854775807)");
    const bytes = sqlite.capi.sqlite3_js_db_export(source.pointer);
    source.close();
    const reader = await openReadOnlySqlite(bytes);
    try { assert.equal(reader.query("SELECT id FROM example").rows[0][0], 9223372036854775807n); }
    finally { reader.close(); }
  `,
    ],
    { cwd: libraryConsumerDirectory, stdio: "inherit" },
  );

  // The whole Power BI path from an installed copy, with no configuration at
  // all: both WebAssembly runtimes have to load themselves, the decoder from
  // the wasm file inside this tarball and SQLite from the peer installed above.
  // Resolving the asset proves it is published; this proves it runs.
  execFileSync(
    process.execPath,
    [
      "--input-type=module",
      "--eval",
      `
    import assert from "node:assert/strict";
    import { readFileSync } from "node:fs";
    import { exportPbiTables } from "@consultchimps/pbi";
    const container = new Uint8Array(readFileSync(${JSON.stringify(pbiFixture)}));
    const outcome = await exportPbiTables(container);
    assert.equal(outcome.outputs.length, 2);
    assert.ok(outcome.outputs[0].bytes.byteLength > 0);
    assert.ok(outcome.result.metrics.exportedTables > 0);
  `,
    ],
    { cwd: libraryConsumerDirectory, stdio: "inherit" },
  );

  // The Power BI package ships its own copy of the third-party notices, because
  // the code they cover is published in its tarball. Prose asking for the two to
  // be kept in step is not a check, so the shared notices are compared here: the
  // headers differ on purpose, the notices themselves may not.
  const noticeBody = (text: string): string => {
    const start = text.indexOf("## XPress9 decoder");
    if (start < 0)
      throw new Error(
        "A THIRD-PARTY-LICENSES.md no longer starts its notices at the XPress9 heading; update scripts/check-packages.ts with the new marker.",
      );
    return text
      .slice(start)
      .replace(/\\r\\n/g, "\n")
      .trimEnd();
  };
  const rootNotices = noticeBody(
    readFileSync(path.join(workspaceRoot, "THIRD-PARTY-LICENSES.md"), "utf8"),
  );
  const packageNotices = noticeBody(
    readFileSync(
      path.join(workspaceRoot, "packages", "pbi", "THIRD-PARTY-LICENSES.md"),
      "utf8",
    ),
  );
  if (rootNotices !== packageNotices) {
    throw new Error(
      "THIRD-PARTY-LICENSES.md at the repository root and in packages/pbi carry different notices. They cover the same vendored code, and the package's copy is the one npm consumers receive, so both must say the same thing.",
    );
  }

  process.stdout.write(
    `Validated ${tarballs.length} package tarballs, the Power BI wasm asset, matching third-party notices, and consultchimps ${installedVersion}.\n`,
  );
} finally {
  rmSync(temporaryRoot, { force: true, recursive: true });
}
