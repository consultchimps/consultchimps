import { execFile } from "node:child_process";
import {
  mkdir,
  mkdtemp,
  readFile,
  readdir,
  realpath,
  rm,
  writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

import JSZip from "jszip";
import { PDFDocument } from "pdf-lib";
import { afterAll, describe, expect, it } from "vitest";

import {
  buildSheetFixture,
  buildWorkbookFixture,
  type FixtureValue,
} from "../../xlsx/test/support/workbook-fixture.js";

// Golden files for the CLI's human-readable output and help. The README beside
// the goldens lists every field the normaliser rewrites and why.

const execFileAsync = promisify(execFile);
const cliPath = fileURLToPath(new URL("../dist/index.js", import.meta.url));
const sourceDirectory = fileURLToPath(new URL("../src/", import.meta.url));
const goldenDirectory = fileURLToPath(new URL("golden/", import.meta.url));
const helpDirectory = path.join(goldenDirectory, "help");
const temporaryDirectories: string[] = [];

afterAll(async () => {
  await Promise.all(
    temporaryDirectories.map((directory) =>
      rm(directory, { force: true, recursive: true }),
    ),
  );
});

// Inherited variables that could change what the CLI prints are dropped, and
// the ones that decide colour, width, time zone, and locale are pinned. Run
// records stay off so nothing is written to the real per-user log folder.
const environment: NodeJS.ProcessEnv = {
  ...Object.fromEntries(
    Object.entries(process.env).filter(
      ([name]) =>
        !/^(?:CONSULTCHIMPS_|FORCE_COLOR$|NODE_OPTIONS$|LC_)/iu.test(name),
    ),
  ),
  CONSULTCHIMPS_LOG: "off",
  COLUMNS: "100",
  // Not C: on Linux, ICU reads C as a POSIX locale that sorts unlike the en-US
  // of a Windows runner, and input discovery sorts with localeCompare.
  LANG: "en_US.UTF-8",
  LC_ALL: "en_US.UTF-8",
  NO_COLOR: "1",
  TZ: "UTC",
};

interface CliRun {
  args: string[];
  exitCode: number;
  stderr: string;
  stdout: string;
}

/** Runs the built CLI with pipes, so it never sees a terminal. */
async function runCli(args: string[], cwd: string): Promise<CliRun> {
  try {
    const result = await execFileAsync(process.execPath, [cliPath, ...args], {
      cwd,
      encoding: "utf8",
      env: environment,
    });
    return { args, exitCode: 0, stderr: result.stderr, stdout: result.stdout };
  } catch (error) {
    const failure = error as Error & {
      code?: number | string;
      stderr?: string;
      stdout?: string;
    };
    if (typeof failure.code !== "number") throw error;
    return {
      args,
      exitCode: failure.code,
      stderr: failure.stderr ?? "",
      stdout: failure.stdout ?? "",
    };
  }
}

/** Escapes text for use inside a pattern. */
function literal(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&");
}

/**
 * Rewrites only what differs between machines: the case's own temporary
 * folder, the Windows path separator, line endings, and terminal control
 * sequences. No covered output prints a duration, date, run id, or version, so
 * none is rewritten; a case that needs one adds it here and to the README.
 */
function normalise(text: string, directories: string[]): string {
  const windows = process.platform === "win32";
  let result = text
    .replace(/\r\n/gu, "\n")
    // eslint-disable-next-line no-control-regex
    .replace(/\u001b(?:\[[0-?]*[ -/]*[@-~]|\][^\u0007]*\u0007)/gu, "");
  for (const directory of directories) {
    // Windows paths compare without regard to case.
    result = result.replace(
      new RegExp(literal(directory), windows ? "giu" : "gu"),
      "<tmp>",
    );
  }
  // Only inside the temporary folder's paths, so a backslash anywhere else
  // still shows up as a difference between platforms.
  return windows
    ? result.replace(/<tmp>[^\s"']*/gu, (found) => found.replace(/\\/gu, "/"))
    : result;
}

function commandLine(args: string[]): string {
  return [
    "consultchimps",
    ...args.map((arg) => (/^[\w./:=@-]+$/u.test(arg) ? arg : `"${arg}"`)),
  ].join(" ");
}

function golden(run: CliRun, directories: string[] = []): string {
  const section = (text: string) =>
    text === "" ? "(empty)\n" : normalise(text, directories);
  return [
    `$ ${commandLine(run.args)}`,
    `exit code: ${run.exitCode}`,
    "",
    "--- stdout ---",
    section(run.stdout),
    "--- stderr ---",
    section(run.stderr),
  ].join("\n");
}

// ---------------------------------------------------------------------------
// Help pages, discovered by walking the help tree of the built CLI.
// ---------------------------------------------------------------------------

/** The subcommand names listed under a help page's Commands heading. */
function listedCommands(help: string): string[] {
  const lines = help.replace(/\r\n/gu, "\n").split("\n");
  const start = lines.indexOf("Commands:");
  if (start === -1) return [];
  const names: string[] = [];
  for (const line of lines.slice(start + 1)) {
    if (line.trim() === "") break;
    const match = /^ {2}(\S+)/u.exec(line);
    if (match && match[1] !== "help") names.push(match[1]!);
  }
  return names;
}

const helpCwd = await mkdtemp(path.join(tmpdir(), "consultchimps-golden-"));
temporaryDirectories.push(helpCwd);
const helpPages = new Map<string, CliRun>();
let level: string[][] = [[]];
try {
  while (level.length > 0) {
    const runs = await Promise.all(
      level.map((names) => runCli([...names, "--help"], helpCwd)),
    );
    level = [];
    for (const run of runs) {
      const names = run.args.slice(0, -1);
      helpPages.set(names.join(" "), run);
      for (const child of listedCommands(run.stdout)) {
        level.push([...names, child]);
      }
    }
  }
} catch (error) {
  // afterAll never runs when the file fails to load.
  await rm(helpCwd, { force: true, recursive: true });
  throw error;
}

function helpGoldenName(command: string): string {
  return `${command === "" ? "consultchimps" : command.replaceAll(" ", ".")}.txt`;
}

describe("help goldens", () => {
  it.concurrent.for([...helpPages.keys()])(
    "consultchimps %s --help",
    async (command, { expect }) => {
      await expect(golden(helpPages.get(command)!)).toMatchFileSnapshot(
        path.join(helpDirectory, helpGoldenName(command)),
      );
    },
  );
});

// ---------------------------------------------------------------------------
// Runs on small neutral fixtures, each in its own temporary folder.
// ---------------------------------------------------------------------------

async function writeWorkbook(
  filePath: string,
  sheets: Array<[string, FixtureValue[][]]>,
): Promise<void> {
  await writeFile(
    filePath,
    await buildWorkbookFixture({
      sheets: sheets.map(([name, rows]) => ({ name, rows })),
    }),
  );
}

const orders = (region: string): FixtureValue[][] => [
  ["Region", "Product", "Amount"],
  [region, "Widgets", 120],
  [region, "Gadgets", 80],
];

async function protectedWorkbook(): Promise<Uint8Array> {
  const archive = await JSZip.loadAsync(
    await buildSheetFixture("Summary", [["Value"], [42]]),
  );
  const workbookXml = await archive.file("xl/workbook.xml")!.async("text");
  archive.file(
    "xl/workbook.xml",
    workbookXml.replace(
      "</workbook>",
      '<workbookProtection lockStructure="1"/></workbook>',
    ),
  );
  const sheetXml = await archive
    .file("xl/worksheets/sheet1.xml")!
    .async("text");
  archive.file(
    "xl/worksheets/sheet1.xml",
    sheetXml.replace(
      "</worksheet>",
      '<sheetProtection sheet="1"/></worksheet>',
    ),
  );
  return archive.generateAsync({ compression: "DEFLATE", type: "uint8array" });
}

async function writePowerPointTemplate(filePath: string): Promise<void> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
  );
  zip.file(
    "_rels/.rels",
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>',
  );
  zip.file(
    "ppt/presentation.xml",
    '<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>',
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>',
  );
  zip.file(
    "ppt/slides/slide1.xml",
    '<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Profile"/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>{{com</a:t></a:r><a:r><a:rPr b="1"/><a:t>pany}}: {{revenue}}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>',
  );
  await writeFile(
    filePath,
    await zip.generateAsync({ compression: "DEFLATE", type: "nodebuffer" }),
  );
}

async function writePdf(filePath: string, pages: number): Promise<void> {
  const document = await PDFDocument.create();
  for (let page = 0; page < pages; page += 1) document.addPage([300, 400]);
  await writeFile(filePath, await document.save());
}

/** Every input any case reads, written once into a case's folder. */
async function writeInputs(directory: string): Promise<void> {
  const inputs = path.join(directory, "inputs");
  await mkdir(inputs);
  await writeWorkbook(path.join(inputs, "north.xlsx"), [
    ["Orders", orders("North")],
  ]);
  await writeWorkbook(path.join(inputs, "south.xlsx"), [
    ["Orders", orders("South")],
  ]);
  await writeFile(
    path.join(inputs, "east.csv"),
    "Region,Product,Amount\nEast,Widgets,95\nEast,Gadgets,60\n",
  );
  await writeFile(
    path.join(inputs, "west.csv"),
    "Region,Product,Amount\nWest,Widgets,70\n",
  );
  await writeWorkbook(path.join(inputs, "orders.xlsx"), [
    [
      "Orders",
      [
        ["Region", "Product", "Amount"],
        ["North", "Widgets", 120],
        ["South", "Gadgets", 80],
        ["North", "Gadgets", 45],
      ],
    ],
  ]);
  await writeFile(
    path.join(inputs, "orders.csv"),
    "Region,Product,Amount\nNorth,Widgets,120\nSouth,Gadgets,80\nNorth,Gadgets,45\n",
  );
  await writeFile(
    path.join(inputs, "protected.xlsx"),
    await protectedWorkbook(),
  );
  await writePowerPointTemplate(path.join(inputs, "profile.pptx"));
  await writeWorkbook(path.join(inputs, "companies.xlsx"), [
    [
      "Companies",
      [
        ["company", "revenue"],
        ["Company A", "12.4"],
        ["Company B", "8.7"],
      ],
    ],
  ]);
  await writeFile(
    path.join(inputs, "companies.csv"),
    "company;revenue\nCompany A;12,4\nCompany B;8,7\n",
  );
  await writePdf(path.join(inputs, "report.pdf"), 3);
  await writePdf(path.join(inputs, "appendix.pdf"), 1);
  await mkdir(path.join(inputs, "chapters"));
  await writePdf(path.join(inputs, "chapters", "two.pdf"), 1);
  await writePdf(path.join(inputs, "chapters", "one.pdf"), 2);
  await writeFile(path.join(inputs, "not-a-pdf.pdf"), "plain text, not a PDF");
  await writeFile(path.join(inputs, "existing.pdf"), "");
}

interface RunCase {
  args: string[];
  name: string;
}

const runCases: RunCase[] = [
  // sheets consolidate
  {
    name: "sheets.consolidate.xlsx",
    args: [
      "sheets",
      "consolidate",
      "inputs/north.xlsx",
      "inputs/south.xlsx",
      "-o",
      "combined.xlsx",
    ],
  },
  {
    name: "sheets.consolidate.csv",
    args: [
      "sheets",
      "consolidate",
      "inputs/east.csv",
      "inputs/west.csv",
      "-o",
      "combined.xlsx",
    ],
  },
  {
    name: "sheets.consolidate.mixed",
    args: [
      "sheets",
      "consolidate",
      "inputs/north.xlsx",
      "inputs/east.csv",
      "-o",
      "combined.xlsx",
    ],
  },
  {
    name: "sheets.consolidate.csv-output",
    args: [
      "sheets",
      "consolidate",
      "inputs/north.xlsx",
      "inputs/east.csv",
      "-o",
      "combined.csv",
    ],
  },
  // sheets merge
  {
    name: "sheets.merge",
    args: [
      "sheets",
      "merge",
      "inputs/north.xlsx",
      "inputs/south.xlsx",
      "inputs/east.csv",
      "-o",
      "merged.xlsx",
    ],
  },
  // sheets split
  {
    name: "sheets.split.column",
    args: [
      "sheets",
      "split",
      "inputs/orders.xlsx",
      "-c",
      "Region",
      "-o",
      "by-region",
    ],
  },
  {
    name: "sheets.split.compact",
    args: [
      "sheets",
      "split",
      "inputs/orders.xlsx",
      "-c",
      "Region",
      "--no-preserve-workbook",
      "-o",
      "by-region",
    ],
  },
  {
    name: "sheets.split.csv",
    args: [
      "sheets",
      "split",
      "inputs/orders.csv",
      "-c",
      "Region",
      "-o",
      "by-region",
    ],
  },
  // sheets inspect
  {
    name: "sheets.inspect.xlsx",
    args: ["sheets", "inspect", "inputs/orders.xlsx"],
  },
  {
    name: "sheets.inspect.csv",
    args: ["sheets", "inspect", "inputs/orders.csv"],
  },
  // sheets unprotect
  {
    name: "sheets.unprotect",
    args: [
      "sheets",
      "unprotect",
      "inputs/protected.xlsx",
      "-o",
      "unprotected.xlsx",
    ],
  },
  // pptx
  {
    name: "pptx.inspect-template",
    args: ["pptx", "inspect-template", "inputs/profile.pptx"],
  },
  {
    name: "pptx.populate",
    args: [
      "pptx",
      "populate",
      "--template",
      "inputs/profile.pptx",
      "--data",
      "inputs/companies.xlsx",
      "-o",
      "profiles.pptx",
    ],
  },
  {
    name: "pptx.populate.csv",
    args: [
      "pptx",
      "populate",
      "--template",
      "inputs/profile.pptx",
      "--data",
      "inputs/companies.csv",
      "--csv-numbers",
      "--csv-decimal",
      ",",
      "-o",
      "profiles.pptx",
    ],
  },
  // pdf
  {
    name: "pdf.split",
    args: ["pdf", "split", "inputs/report.pdf", "-o", "pages"],
  },
  {
    name: "pdf.merge",
    args: [
      "pdf",
      "merge",
      "inputs/report.pdf",
      "inputs/appendix.pdf",
      "-o",
      "combined.pdf",
    ],
  },
  {
    // A named file keeps its place; a folder adds its files alphabetically.
    name: "pdf.merge.folder",
    args: [
      "pdf",
      "merge",
      "inputs/appendix.pdf",
      "inputs/chapters",
      "-o",
      "combined.pdf",
    ],
  },
  // One refusal per command family, and usage errors.
  {
    name: "refusal.sheets.split-missing-column",
    args: [
      "sheets",
      "split",
      "inputs/orders.xlsx",
      "-c",
      "Country",
      "-o",
      "by-country",
    ],
  },
  {
    name: "refusal.pptx.populate-slide-out-of-range",
    args: [
      "pptx",
      "populate",
      "--template",
      "inputs/profile.pptx",
      "--data",
      "inputs/companies.xlsx",
      "--template-slide",
      "4",
      "-o",
      "profiles.pptx",
    ],
  },
  {
    name: "refusal.pptx.populate-csv-sheet",
    args: [
      "pptx",
      "populate",
      "--template",
      "inputs/profile.pptx",
      "--data",
      "inputs/companies.csv",
      "--sheet",
      "Companies",
      "-o",
      "profiles.pptx",
    ],
  },
  {
    name: "refusal.pdf.merge-output-exists",
    args: [
      "pdf",
      "merge",
      "inputs/report.pdf",
      "inputs/appendix.pdf",
      "-o",
      "inputs/existing.pdf",
    ],
  },
  {
    name: "refusal.db.inspect-missing-database",
    args: ["db", "inspect", "missing.duckdb"],
  },
  {
    name: "refusal.usage.missing-output",
    args: ["sheets", "consolidate", "inputs/north.xlsx"],
  },
  {
    name: "refusal.pdf.split-not-a-pdf",
    args: ["pdf", "split", "inputs/not-a-pdf.pdf", "-o", "pages"],
  },
  {
    name: "refusal.usage.header-row-zero",
    args: [
      "sheets",
      "consolidate",
      "inputs/north.xlsx",
      "--header-row",
      "0",
      "-o",
      "combined.xlsx",
    ],
  },
  {
    name: "refusal.usage.split-two-destinations",
    args: [
      "sheets",
      "split",
      "inputs/orders.xlsx",
      "-c",
      "Region",
      "-o",
      "first",
      "--output-dir",
      "second",
    ],
  },
];

describe("run goldens", () => {
  it.concurrent.for(runCases)("$name", async ({ args, name }, { expect }) => {
    const created = await mkdtemp(path.join(tmpdir(), "consultchimps-golden-"));
    temporaryDirectories.push(created);
    await writeInputs(created);
    const run = await runCli(args, created);
    // Every detail is a plain-language label, never a raw metric name.
    expect(run.stdout).not.toMatch(/^ {2}- [a-z]+[A-Z]\w*:/mu);
    // The CLI may print the folder as created, as resolved through links, or
    // (on Windows) in its long form, so all of them are rewritten.
    const directories = [...new Set([await realpath(created), created])].sort(
      (left, right) => right.length - left.length,
    );
    await expect(golden(run, directories)).toMatchFileSnapshot(
      path.join(goldenDirectory, "run", `${name}.txt`),
    );
  });
});

// Runs after both suites. Vitest writes new goldens only as the file ends, so
// after adding a command or case, run `-u` twice (see the README).
describe("golden coverage", () => {
  it("has one help golden for every command and no others", async () => {
    const expected = [...helpPages.keys()].map(helpGoldenName).sort();
    const present = (await readdir(helpDirectory).catch(() => []))
      .filter((name) => name.endsWith(".txt"))
      .sort();
    expect(present).toEqual(expected);
  });

  it("finds every registered command in the help tree", async () => {
    // Names repeat across groups (sheets merge, pdf merge), so each name is
    // counted: a hidden pdf merge must not pass because sheets merge is listed.
    const count = (names: string[]) => {
      const counts = new Map<string, number>();
      for (const name of names) counts.set(name, (counts.get(name) ?? 0) + 1);
      return Object.fromEntries([...counts].sort());
    };
    const listed = [...helpPages.keys()]
      .filter((command) => command !== "")
      .map((command) => command.split(" ").at(-1)!);
    const registered: string[] = [];
    for (const entry of await readdir(sourceDirectory, {
      recursive: true,
    })) {
      if (!entry.endsWith(".ts") || entry.endsWith(".test.ts")) continue;
      const source = await readFile(path.join(sourceDirectory, entry), "utf8");
      for (const match of source.matchAll(/\.command\(\s*"([^"\s]+)/gu)) {
        registered.push(match[1]!);
      }
      // A command built apart and attached, or a hidden one, would dodge it.
      expect(source, entry).not.toMatch(
        /addCommand\(|hidden:\s*true|hideHelp\(/u,
      );
    }
    expect(registered.length).toBeGreaterThan(10);
    expect(count(listed)).toEqual(count(registered));
  });

  it("has one run golden for every case and no others", async () => {
    const expected = runCases.map(({ name }) => `${name}.txt`).sort();
    const present = (
      await readdir(path.join(goldenDirectory, "run")).catch(() => [])
    )
      .filter((name) => name.endsWith(".txt"))
      .sort();
    expect(present).toEqual(expected);
  });
});
