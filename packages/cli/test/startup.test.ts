import { execFile } from "node:child_process";
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { promisify } from "node:util";

import { afterAll, beforeAll, describe, expect, it } from "vitest";

const execFileAsync = promisify(execFile);
const packageDirectory = fileURLToPath(new URL("..", import.meta.url));
const cliPath = path.join(packageDirectory, "dist", "index.js");

// The packages a command loads on demand (src/modules.ts). Loading any of them,
// or the native engines behind the database entry, at startup brings back the
// multi-second delay every command used to pay.
const HEAVY =
  /\/packages\/(xlsx|pdf|pptx|db|files)\/|better-sqlite3|@duckdb|node-api/u;
// The portable bundle inlines every package, so only the external native
// engines can show up as separate modules.
const NATIVE = /better-sqlite3|@duckdb/u;

let directory: string;
let registerUrl: string;

beforeAll(async () => {
  directory = await mkdtemp(path.join(tmpdir(), "consultchimps-startup-"));
  // module.register (Node 20.6+) rather than the synchronous registerHooks,
  // which the supported Node 22.14 floor does not have. The hook runs on its
  // own thread and appends every loaded module URL to the file named in
  // CONSULTCHIMPS_LOAD_LOG. Hooks registered this way do not see require()
  // on that floor, so the preload also records every native addon passed to
  // process.dlopen, which both engines go through however they are loaded.
  const hooks = path.join(directory, "hooks.mjs");
  await writeFile(
    hooks,
    `import { appendFileSync } from "node:fs";
export async function load(url, context, next) {
  appendFileSync(process.env.CONSULTCHIMPS_LOAD_LOG, url + "\\n");
  return next(url, context);
}
`,
  );
  const register = path.join(directory, "register.mjs");
  await writeFile(
    register,
    `import { appendFileSync } from "node:fs";
import { register } from "node:module";
import { pathToFileURL } from "node:url";
register(${JSON.stringify(pathToFileURL(hooks).href)});
const dlopen = process.dlopen;
process.dlopen = function (module, filename, ...rest) {
  appendFileSync(
    process.env.CONSULTCHIMPS_LOAD_LOG,
    pathToFileURL(filename).href + "\\n",
  );
  return dlopen.call(this, module, filename, ...rest);
};
`,
  );
  registerUrl = pathToFileURL(register).href;
});

afterAll(async () => {
  await rm(directory, { force: true, recursive: true });
});

async function loadedModules(
  name: string,
  args: string[],
  entry = cliPath,
): Promise<string[]> {
  const log = path.join(directory, `${name}.log`);
  await writeFile(log, "");
  await execFileAsync(
    process.execPath,
    ["--import", registerUrl, entry, ...args],
    {
      encoding: "utf8",
      env: { ...process.env, CONSULTCHIMPS_LOAD_LOG: log },
    },
  ).catch(() => undefined);
  return (await readFile(log, "utf8"))
    .split("\n")
    .filter((line) => line !== "")
    .map((url) => decodeURIComponent(url).replaceAll("\\", "/"));
}

describe("CLI startup", () => {
  it("loads no heavy package for --version or --help", async () => {
    for (const args of [["--version"], ["--help"], ["sheets", "--help"]]) {
      const loaded = await loadedModules(args.join("-"), args);
      expect(loaded.length, args.join(" ")).toBeGreaterThan(0);
      expect(
        loaded.filter((url) => HEAVY.test(url)),
        args.join(" "),
      ).toEqual([]);
    }
  });

  it("loads a package when a command needs it", async () => {
    // A positive control: without it, a hook that saw nothing would pass the
    // test above.
    const loaded = await loadedModules("pdf-split", [
      "--json",
      "--no-log",
      "pdf",
      "split",
      path.join(directory, "missing.pdf"),
      "-o",
      path.join(directory, "pages"),
    ]);
    expect(loaded.some((url) => url.includes("/packages/files/"))).toBe(true);
  });
});

describe("portable bundle startup", () => {
  let bundleDirectory: string | undefined;
  let bundlePath: string;

  beforeAll(async () => {
    // Inside the package's node_modules so the bundle resolves its external
    // native engines from there, as the portable archive does from its own.
    const cache = path.join(packageDirectory, "node_modules", ".cache");
    await mkdir(cache, { recursive: true });
    bundleDirectory = await mkdtemp(path.join(cache, "startup-bundle-"));
    const tsup = path.join(
      path.dirname(createRequire(import.meta.url).resolve("tsup/package.json")),
      "dist",
      "cli-default.js",
    );
    await execFileAsync(
      process.execPath,
      [
        tsup,
        "--config",
        "tsup.bundle.config.ts",
        "--out-dir",
        bundleDirectory,
        "--silent",
      ],
      { cwd: packageDirectory },
    );
    bundlePath = path.join(bundleDirectory, "consultchimps.mjs");
  }, 600_000);

  afterAll(async () => {
    if (bundleDirectory !== undefined)
      await rm(bundleDirectory, { force: true, recursive: true });
  });

  it("loads no native engine for --version or --help", async () => {
    for (const args of [["--version"], ["--help"], ["db", "--help"]]) {
      const loaded = await loadedModules(
        `bundle-${args.join("-")}`,
        args,
        bundlePath,
      );
      expect(
        loaded.some((url) => url.endsWith("/consultchimps.mjs")),
        args.join(" "),
      ).toBe(true);
      expect(
        loaded.filter((url) => NATIVE.test(url)),
        args.join(" "),
      ).toEqual([]);
    }
  });

  it("loads each native engine when a database opens", async () => {
    // A positive control for both the require and the import path.
    for (const [format, engine] of [
      ["sqlite", "better-sqlite3"],
      ["duckdb", "@duckdb"],
    ] as const) {
      const loaded = await loadedModules(
        `bundle-db-create-${format}`,
        [
          "--json",
          "--no-log",
          "db",
          "create",
          "--format",
          format,
          "-o",
          path.join(directory, `created.${format}`),
        ],
        bundlePath,
      );
      expect(
        loaded.some((url) => url.includes(engine)),
        format,
      ).toBe(true);
    }
  });
});
