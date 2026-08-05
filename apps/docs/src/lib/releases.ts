import { existsSync, readdirSync, readFileSync } from "node:fs";
import path from "node:path";

/**
 * Build-time readers for package versions and changelogs. The docs site is a
 * static export, so everything here runs while `next build` executes inside
 * the monorepo — the deployed pages are plain HTML that updates on every
 * deploy without manual edits.
 */
export interface ReleaseEntry {
  readonly version: string;
  readonly notes: readonly string[];
}

export interface PackageReleases {
  readonly name: string;
  /** Folder under packages/, for building changelog links. */
  readonly folder: string;
  readonly version: string;
  readonly description: string;
  readonly entries: readonly ReleaseEntry[];
}

function repoRoot(): string {
  let dir = process.cwd();
  while (!existsSync(path.join(dir, "pnpm-workspace.yaml"))) {
    const parent = path.dirname(dir);
    if (parent === dir) {
      throw new Error("Could not locate the workspace root from " + dir);
    }
    dir = parent;
  }
  return dir;
}

/** The version consultants actually install; shown in the landing hero. */
export function cliVersion(): string {
  const manifestPath = path.join(repoRoot(), "packages", "cli", "package.json");
  const manifest = JSON.parse(readFileSync(manifestPath, "utf8")) as {
    version: string;
  };
  return manifest.version;
}

const NOTE_PATTERN = /^- (?:[0-9a-f]{7,}: )?(.+)$/;

function parseChangelog(markdown: string, limit: number): ReleaseEntry[] {
  const entries: ReleaseEntry[] = [];
  let current: { version: string; notes: string[] } | undefined;
  // True while the previous line belonged to a kept bullet, so wrapped
  // continuation lines can be folded back into it.
  let collecting = false;

  for (const line of markdown.split(/\r?\n/)) {
    const headingVersion = /^## (\d+\.\d+\.\d+.*)$/.exec(line)?.[1];
    if (headingVersion !== undefined) {
      if (current) entries.push(current);
      if (entries.length >= limit) return entries;
      current = { version: headingVersion, notes: [] };
      collecting = false;
      continue;
    }
    if (!current) continue;

    if (line.startsWith("- ")) {
      const noteText = NOTE_PATTERN.exec(line)?.[1];
      if (
        noteText !== undefined &&
        !noteText.startsWith("Updated dependencies")
      ) {
        current.notes.push(noteText);
        collecting = true;
      } else {
        collecting = false;
      }
      continue;
    }

    const continuation = line.trim();
    if (
      collecting &&
      continuation !== "" &&
      line.startsWith("  ") &&
      !continuation.startsWith("- ")
    ) {
      const previous = current.notes.pop();
      if (previous !== undefined) {
        current.notes.push(`${previous} ${continuation}`);
      }
    } else {
      collecting = false;
    }
  }
  if (current && entries.length < limit) entries.push(current);
  return entries;
}

export function packageReleases(entryLimit = 3): PackageReleases[] {
  const packagesDir = path.join(repoRoot(), "packages");
  const releases: PackageReleases[] = [];

  for (const folder of readdirSync(packagesDir).sort()) {
    // Folder names become URL path segments on /releases; accept only the
    // plain lowercase names our packages actually use.
    if (!/^[a-z0-9-]+$/.test(folder)) continue;
    const manifestPath = path.join(packagesDir, folder, "package.json");
    if (!existsSync(manifestPath)) continue;
    const manifest = JSON.parse(readFileSync(manifestPath, "utf8")) as {
      name: string;
      version: string;
      description?: string;
      private?: boolean;
    };
    if (manifest.private) continue;

    const changelogPath = path.join(packagesDir, folder, "CHANGELOG.md");
    const entries = existsSync(changelogPath)
      ? parseChangelog(readFileSync(changelogPath, "utf8"), entryLimit)
      : [];

    releases.push({
      name: manifest.name,
      folder,
      version: manifest.version,
      description: manifest.description ?? "",
      entries,
    });
  }

  return releases;
}
