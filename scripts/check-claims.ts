import { existsSync, readFileSync, readdirSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import ts from "typescript";

// Claims check: a lint over the copy a user actually reads, guarding the one
// failure mode documentation reviews keep finding here — a sentence that
// promises more than the implementation delivers. Each banned phrase below was
// a real false claim in this repository ("every operation also runs in your
// browser" while the registry declared a planned browser surface; "byte-for-
// byte" for outputs whose filenames deliberately differ; "a complete copy of
// the original" for a split that strips pivot caches by contract).
//
// The rule is not that these words are forbidden English. It is that an
// unconditional claim in user-facing copy has to be held up by something, and
// none of these were. Rewrite the sentence to say what is true, or add an
// allowlist entry below explaining why this particular occurrence is safe.
//
// Scope is deliberately narrow. Tests, changelogs and architecture records are
// excluded, and comments inside the scanned TypeScript are stripped before
// matching: they describe internals to people reading the code, where
// "byte-for-byte" is a precise and testable statement about a ZIP part rather
// than a promise to a user. What survives stripping in a .ts or .tsx file —
// string literals and JSX text — is what the site actually renders.

interface ClaimRule {
  /** Stable id, used by allowlist entries. */
  readonly id: string;
  /** What the phrase tends to over-promise, and what to write instead. */
  readonly advice: string;
  readonly pattern: RegExp;
  /**
   * Restricts the rule to a subset of the scanned files. Absent means every
   * scanned file is a subject.
   */
  readonly appliesTo?: (label: string) => boolean;
}

interface AllowlistEntry {
  /** Repository-relative path, forward slashes. */
  readonly file: string;
  readonly ruleId: string;
  readonly reason: string;
}

const workspaceRoot = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
);

const isPackageReadme = (label: string): boolean =>
  /^packages\/[^/]+\/README\.md$/.test(label);

const CLAIM_RULES: readonly ClaimRule[] = [
  {
    id: "browser-covers-every-operation",
    pattern: /every operation .{0,80}?browser/i,
    advice:
      'the tool registry can declare an operation whose browser surface is "planned"; say "most operations" and point at the per-operation status instead',
  },
  {
    id: "byte-for-byte",
    pattern: /byte-for-byte/i,
    advice:
      "the browser and the command line share the operation code but not the naming or every interface default; claim deterministic file contents for identical inputs and options instead",
  },
  {
    // "complete workbook" is the same promise worn as a noun: the split's
    // preserve-workbook mode keeps the whole workbook, but not completely.
    id: "complete-copy",
    pattern: /complete (?:copy|workbooks?)/i,
    advice:
      'the preserve-workbook split strips pivot tables and their caches by contract (packages/xlsx/src/contract.ts) and leaves external links unrewritten; describe what is preserved, what is removed, and what needs review, and say "the whole workbook" when you mean the mode rather than the guarantee',
  },
  {
    id: "every-tab-or-setting",
    // The trailing \b keeps "every table" out; the claim being linted is the
    // promise that all of a workbook's tabs or settings survive an edit.
    pattern: /every (?:tab|setting)s?\b/i,
    advice:
      'name the structures actually kept ("sheets, formatting, and supported workbook structure") rather than promising all of them',
  },
  {
    id: "useful-worksheet",
    pattern: /useful worksheet/i,
    advice:
      '"useful" hides the real rule: a worksheet is read when it is visible (or includeHiddenSheets is set), passes the sheets filter, and holds at least one non-blank row under a resolvable header',
  },
  {
    id: "stale-node-version",
    pattern: /Node(?:\.js)?\s+24/i,
    advice:
      'published packages declare engines.node ">=22.0.0"; a package README must state the supported floor, not the version the maintainers develop on',
    appliesTo: isPackageReadme,
  },
  {
    id: "counted-packages",
    pattern: /all (?:six|seven|eight) (?:tarballs|packages)/i,
    advice:
      'a spelled-out count goes stale the moment a package is added; use count-free wording such as "a tarball for every published package"',
  },
  {
    id: "always-reflects",
    pattern: /always reflects/i,
    advice:
      "the docs site and the publish workflow run independently and can diverge; say what generates the page and when",
  },
];

// Legitimate occurrences, each with the reason it is not a false claim. Keep
// this list short: needing many entries means the scope below is wrong, not
// that the rules are.
const ALLOWLIST: readonly AllowlistEntry[] = [];

// Files whose copy a user reads: the repository's front door, the guides, the
// rendered site, and the machine-readable summary the site publishes. Package
// CHANGELOGs and docs/adr are deliberately absent — a changelog entry records
// what a past release claimed, and an architecture record argues about
// internals; neither is copy anyone is asked to act on.
const SINGLE_FILES = [
  "README.md",
  "CONTRIBUTING.md",
  "apps/docs/public/llms.txt",
] as const;
const PACKAGE_README_ROOT = "packages";
const SCANNED_TREES = [
  { root: "apps/docs/content", extensions: [".md", ".mdx"] },
  { root: "apps/docs/src", extensions: [".ts", ".tsx"] },
] as const;

function toRepoLabel(absolutePath: string): string {
  return path.relative(workspaceRoot, absolutePath).split(path.sep).join("/");
}

function readTextFile(absolutePath: string): string {
  return readFileSync(absolutePath, "utf8").replace(/\r\n/g, "\n");
}

/**
 * Blank out the comments in TypeScript source, leaving the string literals and
 * JSX text the site actually renders. Comment characters become spaces and
 * newlines are kept, so reported line numbers still point at the right line.
 *
 * The comment ranges come from the TypeScript parser rather than a hand-rolled
 * scanner. A character-level scanner has to guess what a quote character means,
 * and JSX makes that guess wrong: unquoted element text such as `the operation's
 * own result` carries a bare apostrophe that is not a string delimiter, so a
 * scanner following it would stop stripping comments until the next apostrophe.
 * The parser already knows the difference, and the workspace already depends on
 * it for the JavaScript compiler API.
 */
function withoutSourceComments(source: string, fileName: string): string {
  const sourceFile = ts.createSourceFile(
    fileName,
    source,
    ts.ScriptTarget.Latest,
    /* setParentNodes */ true,
    fileName.endsWith(".tsx") ? ts.ScriptKind.TSX : ts.ScriptKind.TS,
  );

  // Every comment is trivia of exactly one token, so walking the token tree
  // reaches all of them: getChildren, not forEachChild, because comments also
  // sit against punctuation such as the braces of a JSX expression. Both sides
  // are collected because the API splits them - a comment that starts on a new
  // line is leading trivia of the token after it, while one that shares a line
  // with the token before it is that token's trailing trivia, which is exactly
  // the `{/* … */}` case. The seen set keeps a range that several ancestors
  // share from being collected twice.
  const ranges: Array<{ end: number; pos: number }> = [];
  const seen = new Set<number>();
  const collect = (found: readonly ts.CommentRange[] | undefined): void => {
    for (const range of found ?? []) {
      if (!seen.has(range.pos)) {
        seen.add(range.pos);
        ranges.push({ end: range.end, pos: range.pos });
      }
    }
  };
  const visit = (node: ts.Node): void => {
    collect(ts.getLeadingCommentRanges(source, node.getFullStart()));
    collect(ts.getTrailingCommentRanges(source, node.getEnd()));
    for (const child of node.getChildren(sourceFile)) {
      visit(child);
    }
  };
  visit(sourceFile);

  // split("") yields UTF-16 code units, which is the unit the parser's
  // positions are expressed in; the spread operator would yield code points
  // and drift on any character outside the basic plane.
  const characters = source.split("");
  for (const range of ranges) {
    for (let index = range.pos; index < range.end; index += 1) {
      if (characters[index] !== "\n") {
        characters[index] = " ";
      }
    }
  }
  return characters.join("");
}

/**
 * The copy of a scanned file the rules are matched against: TypeScript loses
 * its comments, Markdown and plain text are matched whole.
 */
function readScannableText(absolutePath: string): string {
  const text = readTextFile(absolutePath);
  return /\.tsx?$/.test(absolutePath)
    ? withoutSourceComments(text, absolutePath)
    : text;
}

function listFilesWithExtensions(
  directory: string,
  extensions: readonly string[],
): string[] {
  if (!existsSync(directory)) {
    return [];
  }
  const found: string[] = [];
  for (const entry of readdirSync(directory, { withFileTypes: true })) {
    const entryPath = path.join(directory, entry.name);
    if (entry.isDirectory()) {
      found.push(...listFilesWithExtensions(entryPath, extensions));
    } else if (extensions.some((extension) => entry.name.endsWith(extension))) {
      found.push(entryPath);
    }
  }
  return found;
}

function collectScannedFiles(): string[] {
  const files: string[] = [];
  for (const relative of SINGLE_FILES) {
    const absolute = path.join(workspaceRoot, ...relative.split("/"));
    if (existsSync(absolute)) {
      files.push(absolute);
    }
  }

  const packagesDirectory = path.join(workspaceRoot, PACKAGE_README_ROOT);
  if (existsSync(packagesDirectory)) {
    for (const entry of readdirSync(packagesDirectory, {
      withFileTypes: true,
    })) {
      const readme = path.join(packagesDirectory, entry.name, "README.md");
      if (entry.isDirectory() && existsSync(readme)) {
        files.push(readme);
      }
    }
  }

  for (const tree of SCANNED_TREES) {
    files.push(
      ...listFilesWithExtensions(
        path.join(workspaceRoot, ...tree.root.split("/")),
        tree.extensions,
      ),
    );
  }

  return files.filter((file) => !/CHANGELOG\.md$/i.test(file)).sort();
}

const problems: string[] = [];
const usedAllowlistEntries = new Set<string>();
const allowlistKey = (file: string, ruleId: string): string =>
  `${file}::${ruleId}`;
const allowedKeys = new Set(
  ALLOWLIST.map((entry) => allowlistKey(entry.file, entry.ruleId)),
);

const scannedFiles = collectScannedFiles();

for (const absolutePath of scannedFiles) {
  const label = toRepoLabel(absolutePath);
  const lines = readScannableText(absolutePath).split("\n");

  for (const rule of CLAIM_RULES) {
    if (rule.appliesTo && !rule.appliesTo(label)) {
      continue;
    }
    const key = allowlistKey(label, rule.id);
    lines.forEach((line, index) => {
      const match = rule.pattern.exec(line);
      if (match === null) {
        return;
      }
      if (allowedKeys.has(key)) {
        usedAllowlistEntries.add(key);
        return;
      }
      problems.push(
        `${label}:${index + 1} claims "${match[0]}" (rule ${rule.id})\n    ${rule.advice}`,
      );
    });
  }
}

// A phrase that spans a line break is still a claim, so each scanned file is
// also checked with its own whitespace collapsed. Line numbers are not
// available there, so these are reported per file.
for (const absolutePath of scannedFiles) {
  const label = toRepoLabel(absolutePath);
  const collapsed = readScannableText(absolutePath).replace(/\s+/g, " ");

  for (const rule of CLAIM_RULES) {
    if (rule.appliesTo && !rule.appliesTo(label)) {
      continue;
    }
    const key = allowlistKey(label, rule.id);
    const match = rule.pattern.exec(collapsed);
    if (match === null) {
      continue;
    }
    // Marked here too, so an allowed phrase that only ever appears wrapped
    // across two lines still counts its allowlist entry as used.
    if (allowedKeys.has(key)) {
      usedAllowlistEntries.add(key);
      continue;
    }
    const alreadyReportedOnOneLine = problems.some(
      (problem) =>
        problem.startsWith(`${label}:`) &&
        problem.includes(`(rule ${rule.id})`),
    );
    if (alreadyReportedOnOneLine) {
      continue;
    }
    problems.push(
      `${label} claims "${match[0]}" across a line break (rule ${rule.id})\n    ${rule.advice}`,
    );
  }
}

for (const entry of ALLOWLIST) {
  if (!usedAllowlistEntries.has(allowlistKey(entry.file, entry.ruleId))) {
    problems.push(
      `the claims allowlist entry for ${entry.file} (rule ${entry.ruleId}) matches nothing any more; delete it\n    recorded reason: ${entry.reason}`,
    );
  }
}

// Node requirement check: a package README that names a Node version must name
// the one its own engines field supports, expressed as the floor that field
// actually is. This is separate from the banned-phrase rule above, which only
// catches the one stale version we have already shipped.
interface PackageManifest {
  engines?: { node?: string };
}

const nodeMentionPattern = /\bNode(?:\.js)?\s+v?(\d+)(?:\.\d+)*/gi;
const floorSuffixPattern = /^\s*(?:\+|or later|or newer|or above|and later)\b/i;
const packagesDirectory = path.join(workspaceRoot, PACKAGE_README_ROOT);

if (existsSync(packagesDirectory)) {
  for (const entry of readdirSync(packagesDirectory, { withFileTypes: true })) {
    if (!entry.isDirectory()) {
      continue;
    }
    const readmePath = path.join(packagesDirectory, entry.name, "README.md");
    const manifestPath = path.join(
      packagesDirectory,
      entry.name,
      "package.json",
    );
    if (!existsSync(readmePath) || !existsSync(manifestPath)) {
      continue;
    }

    const manifest = JSON.parse(
      readFileSync(manifestPath, "utf8"),
    ) as PackageManifest;
    const declared = manifest.engines?.node;
    const declaredMajor = /(\d+)/.exec(declared ?? "")?.[1];
    const readmeLabel = toRepoLabel(readmePath);
    const manifestLabel = toRepoLabel(manifestPath);
    const readmeText = readTextFile(readmePath);

    const mentions = [...readmeText.matchAll(nodeMentionPattern)];
    if (mentions.length === 0) {
      continue;
    }
    if (declared === undefined || declaredMajor === undefined) {
      problems.push(
        `${readmeLabel} states a Node.js requirement, but ${manifestLabel} declares no engines.node to check it against`,
      );
      continue;
    }

    for (const mention of mentions) {
      const line = readmeText.slice(0, mention.index).split("\n").length;
      if (mention[1] !== declaredMajor) {
        problems.push(
          `${readmeLabel}:${line} says "${mention[0]}", but ${manifestLabel} declares engines.node "${declared}"\n    state the supported floor, or change engines.node if the support window really moved`,
        );
        continue;
      }
      const remainder = readmeText.slice(mention.index + mention[0].length);
      if (!floorSuffixPattern.test(remainder)) {
        problems.push(
          `${readmeLabel}:${line} says "${mention[0]}" without saying it is a floor, but ${manifestLabel} declares engines.node "${declared}"\n    write "Node.js ${declaredMajor} or later" so a reader on a newer Node is not turned away`,
        );
      }
    }
  }
}

if (problems.length > 0) {
  throw new Error(
    `User-facing copy makes claims the implementation does not support:\n${problems
      .map((problem) => `- ${problem}`)
      .join(
        "\n",
      )}\nRewrite each sentence to say what is true, or add an allowlist entry in scripts/check-claims.ts with the reason it is safe.`,
  );
}

process.stdout.write(
  `Checked ${scannedFiles.length} user-facing files against ${CLAIM_RULES.length} claim rules and every package README's Node requirement.\n`,
);
