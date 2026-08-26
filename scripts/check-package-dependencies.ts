import { existsSync, readFileSync, readdirSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

// Published dependency check: every dependency a publishable package declares
// must be installable from the npm registry alone. A tarball URL, git remote,
// or local path works in this workspace but turns a consumer's `npm install`
// into a request to some other host, which fails outright in registry-only,
// mirrored, or air-gapped environments. Code that cannot be depended on from
// the registry has to be bundled at build time instead, which is why
// devDependencies are exempt: they never reach a consumer's install graph.

interface PackageManifest {
  name?: string;
  private?: boolean;
  dependencies?: Record<string, string>;
  peerDependencies?: Record<string, string>;
  optionalDependencies?: Record<string, string>;
}

const workspaceRoot = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
);
// Mirrors the globs in pnpm-workspace.yaml so every workspace project is a
// subject of this check, not just the ones under packages/. An app is private
// today, but a publishable one added later must obey the same rule.
const workspaceDirectoryNames = ["apps", "packages"] as const;

// peerDependencies and optionalDependencies are checked alongside dependencies
// because both end up in a consumer's install graph: npm installs peers
// automatically, and optional entries are installed by default and only
// tolerated when the install itself fails. devDependencies are deliberately
// absent - packages/cli and packages/pptx legitimately point a devDependency at
// a CDN tarball for tests, and nothing there is published.
const publishedDependencyFields = [
  "dependencies",
  "peerDependencies",
  "optionalDependencies",
] as const;

const tarballUrlReason =
  "installs a tarball from a URL outside the npm registry, so a consumer's `npm install` must reach that host";
const gitHostReason =
  "installs from a git host rather than the npm registry, so a consumer's `npm install` must reach that host";
const localPathReason =
  "points at a path that only exists in this workspace, so a consumer's `npm install` has nothing to resolve";
const fixAdvice =
  "bundle the code at build time by moving the entry to devDependencies, or depend on a version published to the npm registry";

// Why each non-registry protocol breaks a consumer's install. Protocols not
// listed here get a generic message; the accepted ones are handled before this
// map is consulted.
const nonRegistryProtocolReasons: ReadonlyMap<string, string> = new Map([
  ["http", tarballUrlReason],
  ["https", tarballUrlReason],
  ["git", gitHostReason],
  ["git+http", gitHostReason],
  ["git+https", gitHostReason],
  ["git+ssh", gitHostReason],
  ["git+file", gitHostReason],
  ["github", gitHostReason],
  ["gist", gitHostReason],
  ["bitbucket", gitHostReason],
  ["gitlab", gitHostReason],
  ["file", localPathReason],
  ["link", localPathReason],
  ["portal", localPathReason],
]);

function toRepoLabel(absolutePath: string): string {
  return path.relative(workspaceRoot, absolutePath).split(path.sep).join("/");
}

// Returns why a specifier is not registry-installable, or null when it is.
// Anything without a protocol and without a path separator is treated as a
// registry range or dist-tag: npm resolves all of those from the registry, and
// range syntax itself is npm's business to validate, not this guardrail's.
function describeNonRegistrySpecifier(specifier: string): string | null {
  const trimmed = specifier.trim();
  const protocolMatch = /^([A-Za-z][A-Za-z0-9+.-]*):/.exec(trimmed);

  if (protocolMatch?.[1] !== undefined) {
    const protocol = protocolMatch[1].toLowerCase();
    // pnpm rewrites `workspace:` and `catalog:` specifiers to the real
    // published version as it packs, so neither form ever reaches a consumer's
    // manifest. `catalog:` is accepted ahead of the workspace adopting
    // catalogs so that adopting them later does not trip this guardrail.
    if (protocol === "workspace" || protocol === "catalog") {
      return null;
    }
    // `npm:<name>` and `npm:<name>@<specifier>` alias one registry package to
    // another name; only the trailing specifier can drag in a non-registry
    // source. The last `@` is the version separator - a leading `@` belongs to
    // a scope, so an alias with no version separator is already acceptable.
    if (protocol === "npm") {
      const aliasTarget = trimmed.slice("npm:".length);
      const versionSeparator = aliasTarget.lastIndexOf("@");
      return versionSeparator > 0
        ? describeNonRegistrySpecifier(aliasTarget.slice(versionSeparator + 1))
        : null;
    }
    return (
      nonRegistryProtocolReasons.get(protocol) ??
      `uses the "${protocol}:" protocol rather than a registry version, so a consumer's \`npm install\` cannot resolve it`
    );
  }

  if (trimmed.startsWith(".")) {
    return localPathReason;
  }
  // A protocol-less specifier containing a slash is npm's `owner/repo`
  // shorthand for a git host; scoped names appear as the dependency key, never
  // here.
  if (trimmed.includes("/")) {
    return "is an owner/repo shorthand for a git host rather than a registry version, so a consumer's `npm install` must reach that host";
  }

  return null;
}

const problems: string[] = [];
let checkedManifestCount = 0;
let checkedSpecifierCount = 0;

// The project list comes from the directory listing rather than a hardcoded
// array so a newly added package is covered the moment it exists.
const manifestPaths = workspaceDirectoryNames
  .flatMap((workspaceDirectoryName) => {
    const workspaceDirectory = path.join(workspaceRoot, workspaceDirectoryName);
    if (!existsSync(workspaceDirectory)) {
      return [];
    }
    return readdirSync(workspaceDirectory, { withFileTypes: true })
      .filter((entry) => entry.isDirectory())
      .map((entry) =>
        path.join(workspaceDirectory, entry.name, "package.json"),
      );
  })
  // A directory without a manifest is not a project - build output, a stray
  // fixture directory - and is simply not a subject of this check.
  .filter((manifestPath) => existsSync(manifestPath))
  .sort();

for (const manifestPath of manifestPaths) {
  const manifest = JSON.parse(
    readFileSync(manifestPath, "utf8"),
  ) as PackageManifest;
  // Private packages are never published, so their install graph is this
  // workspace's problem alone.
  if (manifest.private === true) {
    continue;
  }

  checkedManifestCount += 1;
  const manifestLabel = toRepoLabel(manifestPath);
  const packageLabel =
    manifest.name === undefined
      ? manifestLabel
      : `${manifest.name} (${manifestLabel})`;

  for (const field of publishedDependencyFields) {
    for (const [dependencyName, specifier] of Object.entries(
      manifest[field] ?? {},
    )) {
      checkedSpecifierCount += 1;
      const reason = describeNonRegistrySpecifier(specifier);
      if (reason !== null) {
        problems.push(
          `${packageLabel} declares ${field} entry "${dependencyName}": "${specifier}", which ${reason} and fails in registry-only environments; ${fixAdvice}`,
        );
      }
    }
  }
}

if (problems.length > 0) {
  throw new Error(
    `Publishable package manifests declare dependencies that consumers cannot install from the npm registry:\n${problems
      .map((problem) => `- ${problem}`)
      .join(
        "\n",
      )}\nFix each item so every published dependency resolves from the registry alone.`,
  );
}

process.stdout.write(
  `Verified ${checkedManifestCount} publishable package manifests declare only registry-installable dependencies (${checkedSpecifierCount} entries checked).\n`,
);
