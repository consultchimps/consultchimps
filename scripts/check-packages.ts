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
  "messages",
  "pdf",
  "xlsx",
  "pptx",
  "cli",
] as const;
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

  execFileSync(
    pnpmCommand,
    [
      ...pnpmArguments,
      "--filter",
      "./packages/*",
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

  process.stdout.write(
    `Validated ${tarballs.length} package tarballs and consultchimps ${installedVersion}.\n`,
  );
} finally {
  rmSync(temporaryRoot, { force: true, recursive: true });
}
