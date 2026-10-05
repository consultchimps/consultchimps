// Moves every skill pin, and the version on the agent skills docs page, to a
// newly published CLI release and lists the lines that still name the old
// version, for a human to re-verify.
//
// Run by the post-release job in publish.yml once the release exists, then
// followed by `pnpm skills:reference`, which regenerates the references from
// the new release. Pins move mechanically; a behaviour claim such as "In 0.13.0
// a title block is skipped" does not, because the new release may have changed
// that behaviour. Those lines keep the old version and come out as a checklist
// in the pull request body, so a person re-checks each one against the new
// release before moving it.
//
//   node scripts/update-skills-cli-version.ts 0.14.0 [--body <file>]
import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { parseArgs } from "node:util";

import { format, getFileInfo, resolveConfig } from "prettier";

import {
  DOCS_PAGE,
  DOCS_PIN,
  isGeneratedReference,
  METADATA_PIN,
  pinnedCliVersion,
  RELEASE_VERSION,
  skillDirectories,
  skillTextFiles,
  VERSION_PIN,
  workspaceLabel,
} from "./skill-pins.ts";

const workspaceRoot = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
);

const { values: options, positionals } = parseArgs({
  allowPositionals: true,
  options: { body: { type: "string" } },
});
const next = positionals[0];
if (positionals.length !== 1 || next === undefined) {
  throw new Error(
    "Usage: node scripts/update-skills-cli-version.ts <version> [--body <file>]",
  );
}
if (!RELEASE_VERSION.test(next)) {
  throw new Error(`"${next}" is not a release version such as "1.2.3".`);
}

function compareVersions(a: string, b: string): number {
  const left = a.split(".").map(Number);
  const right = b.split(".").map(Number);
  for (let index = 0; index < 3; index += 1) {
    const difference = (left[index] ?? 0) - (right[index] ?? 0);
    if (difference !== 0) {
      return difference;
    }
  }
  return 0;
}

const previous = pinnedCliVersion(workspaceRoot);
// A patch release on an older line must not move the skills backwards.
if (compareVersions(next, previous) <= 0) {
  process.stdout.write(
    `The skills pin ${previous}, which is not older than ${next}; nothing to move.\n`,
  );
  process.exit(0);
}

const escaped = previous.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
const mentionsPrevious = new RegExp(`(?<![\\d.])${escaped}(?!\\.?\\d)`);
const claims: string[] = [];
let moved = 0;

// A longer version can push a wrapped Markdown line past the print width, and
// the formatting check would then fail the pull request.
async function formatAndWrite(file: string, text: string): Promise<void> {
  const info = await getFileInfo(file, {
    ignorePath: path.join(workspaceRoot, ".prettierignore"),
  });
  let output = text;
  if (!info.ignored && info.inferredParser !== null) {
    const config = (await resolveConfig(file)) ?? {};
    output = await format(text, { ...config, filepath: file });
  }
  writeFileSync(file, output, "utf8");
}

function listClaims(label: string, file: string): void {
  readFileSync(file, "utf8")
    .split(/\r?\n/)
    .forEach((line, index) => {
      if (mentionsPrevious.test(line)) {
        claims.push(`- [ ] \`${label}:${String(index + 1)}\`: ${line.trim()}`);
      }
    });
}

for (const directory of skillDirectories(workspaceRoot)) {
  for (const file of skillTextFiles(directory)) {
    const label = workspaceLabel(workspaceRoot, file);
    if (isGeneratedReference(label)) {
      continue;
    }
    const original = readFileSync(file, "utf8");
    // pinnedCliVersion has already held every pin to the previous version.
    let text = original.replace(VERSION_PIN, (pin, version: string) => {
      moved += 1;
      return pin.replace(version, next);
    });
    // Only the frontmatter holds the metadata pin, as collectPins reads it.
    const frontmatterEnd = text.indexOf("\n---", 4);
    if (path.basename(file) === "SKILL.md" && frontmatterEnd !== -1) {
      text =
        text
          .slice(0, frontmatterEnd)
          .replace(
            METADATA_PIN,
            (_match, before: string, _version: string, after: string) => {
              moved += 1;
              return `${before}${next}${after}`;
            },
          ) + text.slice(frontmatterEnd);
    }
    if (text !== original) {
      await formatAndWrite(file, text);
    }
    listClaims(label, file);
  }
}

// The docs page names the checked-against version once, and check-skills.ts
// holds it to the pins. Any other mention there is a claim to re-verify.
const docsFile = path.join(workspaceRoot, ...DOCS_PAGE.split("/"));
const docsOriginal = readFileSync(docsFile, "utf8");
const docsText = docsOriginal.replace(
  DOCS_PIN,
  (_match, before: string) => `${before}${next}`,
);
if (docsText !== docsOriginal) {
  moved += 1;
  await formatAndWrite(docsFile, docsText);
}
listClaims(DOCS_PAGE, docsFile);

const body = [
  `Moves every skill pin from consultchimps ${previous} to ${next} and regenerates both CLI references from the ${next} release.`,
  "",
  "## Claims to re-verify",
  "",
  claims.length === 0
    ? `No line in \`skills/\` or \`${DOCS_PAGE}\` names ${previous} outside the pins.`
    : `These lines still name ${previous}. Check each against ${next}, then move it to ${next} or rewrite it, in this pull request.\n\n${claims.join("\n")}`,
  "",
].join("\n");
if (options.body === undefined) {
  process.stdout.write(body);
} else {
  writeFileSync(options.body, body, "utf8");
}
process.stdout.write(
  `Moved ${String(moved)} pins from ${previous} to ${next}; ${String(claims.length)} lines still name ${previous}.\n`,
);
