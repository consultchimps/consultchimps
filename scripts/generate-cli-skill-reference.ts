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
//
// The skills run the published CLI (`npx consultchimps@X`), so the reference
// is read from the published release of the version the skills pin, the
// single-file `consultchimps.mjs` attached to its GitHub release, never from
// packages/cli/package.json or the checkout: between releases the checkout can
// print help the published version does not have, and a release pull request
// bumps the package version before that version exists. The file is downloaded
// afresh into a private temporary directory and checked against the SHA-256
// digest GitHub publishes for the release asset before it runs; nothing is
// reused from a shared path. A version with no release, or any other failure
// such as GitHub being unreachable, stops the run.
//
// `--version X` writes the references for another release (the post-release
// update moves the pins first, so the default is usually right). `--local`
// reads the local build instead, for developing the generator; its output is
// labelled with the target version and the check will reject it until that
// version publishes with the same help.
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { parseArgs } from "node:util";

import { pinnedCliVersion, RELEASE_VERSION } from "./skill-pins.ts";

const workspaceRoot = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
);
const localCliPath = path.join(
  workspaceRoot,
  "packages",
  "cli",
  "dist",
  "index.js",
);
const generatorLabel = "pnpm skills:reference";

const repository = "consultchimps/consultchimps";

interface ReleaseAsset {
  name?: unknown;
  digest?: unknown;
  browser_download_url?: unknown;
}

/**
 * The published single-file CLI for `version`, verified against its release
 * digest, in a fresh private directory; or null when that version has no
 * release yet. Every other failure throws.
 */
async function releasedCli(
  version: string,
): Promise<{ file: string; directory: string } | null> {
  const headers: Record<string, string> = {
    accept: "application/vnd.github+json",
  };
  const token = process.env.GITHUB_TOKEN;
  if (token) {
    headers.authorization = `Bearer ${token}`;
  }
  const tag = encodeURIComponent(`consultchimps@${version}`);
  const releaseResponse = await fetch(
    `https://api.github.com/repos/${repository}/releases/tags/${tag}`,
    { headers },
  );
  if (releaseResponse.status === 404) {
    return null;
  }
  if (!releaseResponse.ok) {
    throw new Error(
      `Could not read the consultchimps ${version} release from GitHub (HTTP ${String(releaseResponse.status)}). The reference is generated from that release; retry when GitHub is reachable.`,
    );
  }
  const release = (await releaseResponse.json()) as { assets?: unknown };
  const asset = (Array.isArray(release.assets) ? release.assets : []).find(
    (candidate: ReleaseAsset) => candidate.name === "consultchimps.mjs",
  ) as ReleaseAsset | undefined;
  if (
    typeof asset?.browser_download_url !== "string" ||
    typeof asset.digest !== "string" ||
    !asset.digest.startsWith("sha256:")
  ) {
    throw new Error(
      `The consultchimps ${version} release has no consultchimps.mjs with a SHA-256 digest to verify.`,
    );
  }
  const download = await fetch(asset.browser_download_url);
  if (!download.ok) {
    throw new Error(
      `Could not download consultchimps.mjs for ${version} (HTTP ${String(download.status)}).`,
    );
  }
  const bytes = Buffer.from(await download.arrayBuffer());
  const digest = createHash("sha256").update(bytes).digest("hex");
  if (`sha256:${digest}` !== asset.digest) {
    throw new Error(
      `consultchimps.mjs for ${version} does not match its published digest; refusing to run it.`,
    );
  }
  const directory = mkdtempSync(path.join(tmpdir(), "consultchimps-release-"));
  const file = path.join(directory, "consultchimps.mjs");
  writeFileSync(file, bytes, { mode: 0o600 });
  return { file, directory };
}

const { values: options } = parseArgs({
  options: {
    check: { type: "boolean", default: false },
    local: { type: "boolean", default: false },
    version: { type: "string" },
  },
});
const cliVersion = options.version ?? pinnedCliVersion(workspaceRoot);
if (!RELEASE_VERSION.test(cliVersion)) {
  throw new Error(`"${cliVersion}" is not a release version such as "1.2.3".`);
}
const released = options.local ? null : await releasedCli(cliVersion);
if (!options.local && released === null) {
  throw new Error(
    `consultchimps ${cliVersion} has no GitHub release. The skills pin published releases only, and move to a new one after it publishes; pass --local to read the local build while developing.`,
  );
}
const cliPath = released?.file ?? localCliPath;
const cliSource =
  released === null
    ? `the local build, labelled ${cliVersion} (development only)`
    : `the published consultchimps ${cliVersion} release`;
process.on("exit", () => {
  if (released !== null) {
    rmSync(released.directory, { recursive: true, force: true });
  }
});

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

// Only what Node needs to start and locate temporary files on each platform.
const helpEnvironment: NodeJS.ProcessEnv = { FORCE_COLOR: "0", NO_COLOR: "1" };
for (const name of ["PATH", "Path", "SystemRoot", "TEMP", "TMP", "TMPDIR"]) {
  const value = process.env[name];
  if (value !== undefined) {
    helpEnvironment[name] = value;
  }
}

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
      // A minimal environment: the downloaded CLI is verified against its
      // digest, but it still has no reason to see tokens or other secrets.
      env: helpEnvironment,
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
  const commandPaths = discoverCommandPaths([]).filter(target.includes);
  const sections = commandPaths.map((commandPath) => {
    const label = ["consultchimps", ...commandPath].join(" ");
    return `## ${label}\n\n\`\`\`text\n${readHelpText(commandPath)}\n\`\`\``;
  });

  return [
    `<!-- Generated from the CLI's own help by scripts/generate-cli-skill-reference.ts. Do not edit; run \`${generatorLabel}\`. -->`,
    "",
    "# ConsultChimps CLI reference",
    "",
    `The ${target.scope} of \`consultchimps\` ${cliVersion}, as the CLI itself`,
    "prints them. A flag absent here does not exist in that version.",
    "",
    ...(target.skill === "use-consultchimps" &&
    /^ {2}db\s/m.test(readHelpText([]))
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

const check = options.check;

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
        `${referenceLabel} no longer matches the help of ${cliSource}. The skill would document a command surface the CLI does not have. Run \`${generatorLabel}\` and commit the result.`,
      );
    }
    process.stdout.write(
      `Verified ${referenceLabel} against ${String(commandCount(generated))} commands of ${cliSource}.\n`,
    );
  } else {
    writeFileSync(referencePath, generated, "utf8");
    process.stdout.write(
      `Wrote ${referenceLabel} from ${String(commandCount(generated))} commands of ${cliSource}.\n`,
    );
  }
}
