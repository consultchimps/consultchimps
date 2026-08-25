import { spawnSync } from "node:child_process";
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

interface DiscoveredCommand {
  // The subcommand words after the executable name, e.g. ["sheets", "split"].
  // The root program itself is represented by an empty path.
  commandPath: string[];
  // Long option flags declared in the command's Options: section, excluding
  // --help and --version, which Commander adds to every command automatically.
  longOptions: string[];
}

const workspaceRoot = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
);
const cliPath = path.join(workspaceRoot, "packages", "cli", "dist", "index.js");
const referencePagePath = path.join(
  workspaceRoot,
  "apps",
  "docs",
  "content",
  "docs",
  "reference",
  "cli.mdx",
);
// The repository-relative POSIX path keeps error messages stable across
// platforms and matches how the page is referred to elsewhere in the docs.
const referencePageLabel = "apps/docs/content/docs/reference/cli.mdx";

if (!existsSync(cliPath)) {
  throw new Error(
    `The built CLI was not found at ${cliPath}. Run \`pnpm build\` first so the help output can be compared against the CLI reference page.`,
  );
}

if (!existsSync(referencePagePath)) {
  throw new Error(
    `The CLI reference page was not found at ${referencePagePath}.`,
  );
}

// Help is read with stdio piped rather than inherited so Commander sees a
// non-TTY stream and deterministically hard-wraps at its 80-column fallback on
// every OS. NO_COLOR/FORCE_COLOR guard against ANSI escape codes leaking into
// the text that the section parser scans.
function readHelpText(commandPath: string[]): string {
  const commandLabel = ["consultchimps", ...commandPath].join(" ");
  const spawned = spawnSync(
    process.execPath,
    [cliPath, ...commandPath, "--help"],
    {
      encoding: "utf8",
      env: { ...process.env, FORCE_COLOR: "0", NO_COLOR: "1" },
    },
  );

  if (spawned.error) {
    throw new Error(
      `Failed to run "${commandLabel} --help": ${spawned.error.message}`,
    );
  }
  // The CLI uses Commander's exitOverride, but --help is a successful
  // termination whose output is already written, so it still exits 0.
  if (spawned.status !== 0) {
    throw new Error(
      `"${commandLabel} --help" exited with status ${spawned.status}: ${spawned.stderr}`,
    );
  }

  const helpText = spawned.stdout.replace(/\r\n/g, "\n");
  if (helpText.trim() === "") {
    throw new Error(`"${commandLabel} --help" produced no help text.`);
  }
  return helpText;
}

// Walks the CLI's help output recursively and records every command with its
// long options. Extraction is deliberately scoped to Commander's Options: and
// Commands: sections: the addHelpText epilogues contain example command lines
// full of flags, and reading those would add spurious requirements.
function discoverCommands(commandPath: string[]): DiscoveredCommand[] {
  const helpText = readHelpText(commandPath);
  const longOptions: string[] = [];
  const subcommands: string[] = [];
  let currentSection: string | null = null;

  for (const line of helpText.split("\n")) {
    // Commander renders every section header - Usage:, Arguments:, Options:,
    // Commands:, and any epilogue heading such as Examples: - as a
    // non-indented line ending in a colon. Any unrecognized header simply
    // switches the parser into an ignored section.
    if (/^\S.*:$/.test(line)) {
      currentSection = line;
      continue;
    }

    if (currentSection === "Commands:") {
      // Entry rows use exactly two spaces of indentation; wrapped description
      // continuations are indented deeper and must not be read as commands.
      const entry = /^ {2}(\S+)/.exec(line);
      if (entry?.[1] !== undefined && entry[1] !== "help") {
        subcommands.push(entry[1]);
      }
      continue;
    }

    if (currentSection === "Options:" && /^ {2}-/.test(line)) {
      // Only the flags column - the text before the first run of two or more
      // spaces - may contribute flags; the description column can legitimately
      // mention other options, such as "(alias for --output)".
      const flagsColumn = line.trimStart().split(/\s{2,}/, 1)[0] ?? "";
      for (const flag of flagsColumn.match(/--[A-Za-z0-9][A-Za-z0-9-]*/g) ??
        []) {
        if (flag !== "--help" && flag !== "--version") {
          longOptions.push(flag);
        }
      }
    }
  }

  return [
    { commandPath, longOptions },
    ...subcommands.flatMap((subcommand) =>
      discoverCommands([...commandPath, subcommand]),
    ),
  ];
}

const discovered = discoverCommands([]);
const referenceText = readFileSync(referencePagePath, "utf8").replace(
  /\r\n/g,
  "\n",
);
// Command phrases in the docs may be wrapped or re-indented, so they are
// matched against a whitespace-collapsed copy of the page.
const collapsedReferenceText = referenceText.replace(/\s+/g, " ");

const missing: string[] = [];
let verifiedCommandCount = 0;
let verifiedOptionCount = 0;

for (const { commandPath, longOptions } of discovered) {
  const commandLabel = ["consultchimps", ...commandPath].join(" ");

  // The root program has no command phrase of its own to document; only its
  // options - currently --json - need to appear on the reference page.
  if (commandPath.length > 0) {
    verifiedCommandCount += 1;
    if (!collapsedReferenceText.includes(commandLabel)) {
      missing.push(`command "${commandLabel}"`);
    }
  }

  for (const flag of longOptions) {
    verifiedOptionCount += 1;
    // The trailing boundary prevents a longer flag from satisfying a shorter
    // one, e.g. --header-rows must not count as documenting --header-row. Any
    // occurrence in the page - prose or code block - counts as documented.
    const escapedFlag = flag.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    if (!new RegExp(`${escapedFlag}(?![A-Za-z0-9-])`).test(referenceText)) {
      missing.push(`option ${flag} of command "${commandLabel}"`);
    }
  }
}

if (missing.length > 0) {
  throw new Error(
    `The CLI reference page ${referencePageLabel} is missing documentation for:\n${missing
      .map((entry) => `- ${entry}`)
      .join(
        "\n",
      )}\nAdd each missing command and option to the reference page so the documentation matches the CLI.`,
  );
}

process.stdout.write(
  `Verified ${verifiedCommandCount} CLI commands and ${verifiedOptionCount} options against the CLI reference page.\n`,
);
