import { execFile } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { promisify } from "node:util";

import { afterAll, beforeAll, describe, expect, it } from "vitest";

const execFileAsync = promisify(execFile);
const cliPath = fileURLToPath(new URL("../dist/index.js", import.meta.url));

// The packages a command loads on demand (src/modules.ts). Loading any of them,
// or the native engines behind the database entry, at startup brings back the
// multi-second delay every command used to pay.
const HEAVY =
  /\/packages\/(xlsx|pdf|pptx|db|files)\/|better-sqlite3|@duckdb|node-api/u;

let directory: string;
let registerUrl: string;

beforeAll(async () => {
  directory = await mkdtemp(path.join(tmpdir(), "consultchimps-startup-"));
  // module.register (Node 20.6+) rather than the synchronous registerHooks,
  // which the supported Node 22.14 floor does not have. The hook runs on its
  // own thread and appends every loaded module URL to the file named in
  // CONSULTCHIMPS_LOAD_LOG.
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
    `import { register } from "node:module";
register(${JSON.stringify(pathToFileURL(hooks).href)});
`,
  );
  registerUrl = pathToFileURL(register).href;
});

afterAll(async () => {
  await rm(directory, { force: true, recursive: true });
});

async function loadedModules(name: string, args: string[]): Promise<string[]> {
  const log = path.join(directory, `${name}.log`);
  await writeFile(log, "");
  await execFileAsync(
    process.execPath,
    ["--import", registerUrl, cliPath, ...args],
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
