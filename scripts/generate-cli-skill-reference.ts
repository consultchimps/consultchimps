// The CLI references bundled with the tool skills are generated from the built
// CLI, never written by hand, for the same reason
// check-cli-reference.ts exists: a hand-written command list drifts, and a
// skill that names a flag the CLI does not have sends an agent down a path
// that ends in an error it cannot diagnose.
//
// The generator is also the check. `--check` regenerates into memory and
// compares against the committed files, so `pnpm docs:check` fails when the CLI
// gains, loses, or renames anything a skill documents, and the fix is one
// command rather than an edit.
//
// Each tool skill carries its own copy of the commands it drives. A registry
// such as skills.sh installs one skill on its own, so a link from one skill
// into another skill's folder would break after install.
import { spawnSync } from "node:child_process";
import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const workspaceRoot = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
);
const cliPath = path.join(workspaceRoot, "packages", "cli", "dist", "index.js");
const generatorLabel = "pnpm skills:reference";

/**
 * Help is read with stdio piped so Commander sees a non-TTY stream and wraps
 * at its 80-column fallback on every platform, and with colour disabled so no
 * escape codes reach the committed file. Both matter: the file is compared
 * byte for byte, so the same CLI has to print the same bytes on Windows and
 * Linux.
 */
// Discovery and rendering each ask for the same command's help, and every ask
// is a process spawn. Twenty-seven commands is fifty-four spawns without this,
// which is the difference between a fast check and one people skip.
const helpCache = new Map<string, string>();

function readHelpText(commandPath: readonly string[]): string {
  const label = ["consultchimps", ...commandPath].join(" ");
  const cached = helpCache.get(label);
  if (cached !== undefined) {
    return cached;
  }
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
      `Failed to run "${label} --help": ${spawned.error.message}`,
    );
  }
  if (spawned.status !== 0) {
    throw new Error(
      `"${label} --help" exited with status ${String(spawned.status)}: ${spawned.stderr}`,
    );
  }
  const helpText = spawned.stdout.replace(/\r\n/g, "\n").trimEnd();
  if (helpText.trim() === "") {
    throw new Error(`"${label} --help" produced no help text.`);
  }
  helpCache.set(label, helpText);
  return helpText;
}

// The database commands need native bindings, ship no recipes in any skill,
// and are thirteen of the CLI's twenty-seven commands. Documenting them would
// make the reference mostly noise for an agent doing document work, and would
// double what the drift check spawns.
const EXCLUDED_SUBTREES = ["db"];

/**
 * Every command the CLI exposes below the excluded subtrees, parents before
 * children, in the order Commander lists them. The Commands: section is the
 * only source of subcommands: the epilogues carry example command lines, and
 * reading those would invent commands that do not exist.
 */
function discoverCommandPaths(commandPath: readonly string[]): string[][] {
  const subcommands: string[] = [];
  let section: string | null = null;

  for (const line of readHelpText(commandPath).split("\n")) {
    if (/^\S.*:$/.test(line)) {
      section = line;
      continue;
    }
    if (section !== "Commands:") {
      continue;
    }
    // Entry rows carry exactly two spaces of indentation; wrapped description
    // lines are indented deeper and are not commands.
    const entry = /^ {2}(\S+)/.exec(line);
    if (
      entry?.[1] !== undefined &&
      entry[1] !== "help" &&
      !(commandPath.length === 0 && EXCLUDED_SUBTREES.includes(entry[1]))
    ) {
      subcommands.push(entry[1]);
    }
  }

  return [
    [...commandPath],
    ...subcommands.flatMap((subcommand) =>
      discoverCommandPaths([...commandPath, subcommand]),
    ),
  ];
}

function readCliVersion(): string {
  const manifest: unknown = JSON.parse(
    readFileSync(path.join(workspaceRoot, "packages", "cli", "package.json"), {
      encoding: "utf8",
    }),
  );
  if (
    typeof manifest !== "object" ||
    manifest === null ||
    !("version" in manifest) ||
    typeof manifest.version !== "string"
  ) {
    throw new Error("packages/cli/package.json declares no version string.");
  }
  return manifest.version;
}

interface ReferenceTarget {
  /** The skill directory the reference is written into. */
  readonly skill: string;
  /** Which command paths the skill documents. */
  readonly includes: (commandPath: readonly string[]) => boolean;
  /** What the reference covers, completing "The ... of `consultchimps` X". */
  readonly scope: string;
}

const TARGETS: readonly ReferenceTarget[] = [
  {
    skill: "use-consultchimps",
    includes: () => true,
    scope: "document commands",
  },
  {
    skill: "chimps-xlsx",
    includes: (commandPath) => commandPath[0] === "sheets",
    scope: "`sheets` commands",
  },
];

function render(target: ReferenceTarget): string {
  const version = readCliVersion();
  const commandPaths = discoverCommandPaths([]).filter(target.includes);
  const sections = commandPaths.map((commandPath) => {
    const label = ["consultchimps", ...commandPath].join(" ");
    return `## ${label}\n\n\`\`\`text\n${readHelpText(commandPath)}\n\`\`\``;
  });

  return [
    `<!-- Generated from the built CLI by scripts/generate-cli-skill-reference.ts. Do not edit; run \`${generatorLabel}\`. -->`,
    "",
    "# ConsultChimps CLI reference",
    "",
    `The ${target.scope} of \`consultchimps\` ${version}, as the CLI itself`,
    "prints them. A flag absent here does not exist in that version.",
    "",
    ...(target.skill === "use-consultchimps"
      ? [
          `The \`${EXCLUDED_SUBTREES.join("`, `")}\` commands are left out: they need`,
          "native database bindings and no skill carries recipes for them. Run",
          "`consultchimps db --help` against an install to see them.",
          "",
        ]
      : []),
    ...sections.flatMap((section) => [section, ""]),
  ]
    .join("\n")
    .trimEnd()
    .concat("\n");
}

/** How many commands a rendered reference documents, for the summary line. */
function commandCount(reference: string): number {
  return reference.match(/^## /gm)?.length ?? 0;
}

const check = process.argv.includes("--check");

for (const target of TARGETS) {
  const referenceLabel = `skills/${target.skill}/references/cli-reference.md`;
  const referencePath = path.join(workspaceRoot, ...referenceLabel.split("/"));
  const generated = render(target);

  if (check) {
    let committed: string;
    try {
      committed = readFileSync(referencePath, "utf8").replace(/\r\n/g, "\n");
    } catch {
      throw new Error(
        `${referenceLabel} is missing. Run \`${generatorLabel}\` to create it.`,
      );
    }
    if (committed !== generated) {
      throw new Error(
        `${referenceLabel} no longer matches the built CLI's help output. The skill would document a command surface the CLI does not have. Run \`${generatorLabel}\` and commit the result.`,
      );
    }
    process.stdout.write(
      `Verified ${referenceLabel} against ${String(commandCount(generated))} commands of the built CLI.\n`,
    );
  } else {
    writeFileSync(referencePath, generated, "utf8");
    process.stdout.write(
      `Wrote ${referenceLabel} from ${String(commandCount(generated))} commands of the built CLI.\n`,
    );
  }
}
