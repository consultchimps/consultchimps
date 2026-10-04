// The CLI version the skills pin, shared by the skills check, the reference
// generator and the post-release update so all three read pins the same way.
//
// A pin is `metadata.cli-version` in a SKILL.md frontmatter, an
// `npx consultchimps@X` command, or a release download URL
// (`consultchimps%40X`). Every pin across every skill names one version, a
// published CLI release, and the skills move to a new release in a pull
// request of their own after it publishes, never on the release pull request.
import { readFileSync, readdirSync, statSync } from "node:fs";
import path from "node:path";

export const VERSION_PIN: RegExp = /consultchimps(?:@|%40)(\d+\.\d+\.\d+)/g;
export const METADATA_PIN: RegExp =
  /^( {2}cli-version:\s*["']?)(\d+\.\d+\.\d+)(["']?\s*)$/m;
export const RELEASE_VERSION: RegExp = /^\d+\.\d+\.\d+$/;

/** Generated files carry the version they were generated from, not a pin. */
export const GENERATED_REFERENCE = "references/cli-reference.md";

export interface Pin {
  /** Path relative to the workspace root, with forward slashes. */
  readonly file: string;
  readonly line: number;
  readonly version: string;
}

/** Text files in a skill, the only files a pin or a claim can sit in. */
export function skillTextFiles(directory: string): string[] {
  return readdirSync(directory).flatMap((entry) => {
    const full = path.join(directory, entry);
    if (statSync(full).isDirectory()) {
      return skillTextFiles(full);
    }
    return /\.(?:md|py|html|css|js|ts|sh|ps1)$/.test(entry) ? [full] : [];
  });
}

export function skillDirectories(workspaceRoot: string): string[] {
  const skillsRoot = path.join(workspaceRoot, "skills");
  return readdirSync(skillsRoot)
    .map((entry) => path.join(skillsRoot, entry))
    .filter((full) => statSync(full).isDirectory());
}

export function workspaceLabel(workspaceRoot: string, file: string): string {
  return path.relative(workspaceRoot, file).split(path.sep).join("/");
}

/** Whether a skill file is one of the generated references. */
export function isGeneratedReference(label: string): boolean {
  return label.endsWith(`/${GENERATED_REFERENCE}`);
}

function lineOf(text: string, index: number): number {
  return text.slice(0, index).split("\n").length;
}

/** Every pin in every skill, in file order. */
export function collectPins(workspaceRoot: string): Pin[] {
  const pins: Pin[] = [];
  for (const directory of skillDirectories(workspaceRoot)) {
    for (const file of skillTextFiles(directory)) {
      const label = workspaceLabel(workspaceRoot, file);
      if (isGeneratedReference(label)) {
        continue;
      }
      const text = readFileSync(file, "utf8").replace(/\r\n/g, "\n");
      if (path.basename(file) === "SKILL.md") {
        const frontmatterEnd = text.indexOf("\n---", 4);
        const metadata =
          frontmatterEnd === -1
            ? null
            : METADATA_PIN.exec(text.slice(0, frontmatterEnd));
        if (metadata?.[2] !== undefined) {
          pins.push({
            file: label,
            line: lineOf(text, metadata.index),
            version: metadata[2],
          });
        }
      }
      for (const match of text.matchAll(VERSION_PIN)) {
        pins.push({
          file: label,
          line: lineOf(text, match.index),
          version: match[1] ?? "",
        });
      }
    }
  }
  return pins;
}

/**
 * The one version every pin names. Throws when the skills pin nothing or
 * disagree, naming each pin off the majority version.
 */
export function pinnedCliVersion(workspaceRoot: string): string {
  const pins = collectPins(workspaceRoot);
  const counts = new Map<string, number>();
  for (const pin of pins) {
    counts.set(pin.version, (counts.get(pin.version) ?? 0) + 1);
  }
  const ranked = [...counts].sort((a, b) => b[1] - a[1]);
  const version = ranked[0]?.[0];
  if (version === undefined) {
    throw new Error("No skill pins a consultchimps version.");
  }
  if (ranked.length > 1) {
    const strays = pins
      .filter((pin) => pin.version !== version)
      .map((pin) => `  - ${pin.file}:${String(pin.line)} pins ${pin.version}`);
    throw new Error(
      `The skills pin more than one consultchimps version; every pin must name ${version}, the one most of them name:\n${strays.join("\n")}`,
    );
  }
  return version;
}
