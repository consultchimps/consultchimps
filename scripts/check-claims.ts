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
  /**
   * One-based line the excused claim starts on. Together with the phrase it
   * identifies one occurrence, so an exemption can never widen into a blanket
   * pardon: a second "complete copy" further down the same file is still
   * reported, and so is a different phrase under the same rule. A line number
   * moves, but moving it can only make the claim reappear as a failure and the
   * entry report itself stale - never make an unexamined claim pass quietly.
   */
  readonly line: number;
  /** The exact text the rule matches there, compared case insensitively. */
  readonly phrase: string;
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
    pattern: /every operation .{0,80}?browser/gi,
    advice:
      'the tool registry can declare an operation whose browser surface is "planned"; say "most operations" and point at the per-operation status instead',
  },
  {
    id: "byte-for-byte",
    pattern: /byte-for-byte/gi,
    advice:
      "the browser and the command line share the operation code but not the naming or every interface default; claim deterministic file contents for identical inputs and options instead",
  },
  {
    // "complete workbook" is the same promise worn as a noun: the split's
    // preserve-workbook mode keeps the whole workbook, but not completely.
    id: "complete-copy",
    pattern: /complete (?:copy|workbooks?)/gi,
    advice:
      'the preserve-workbook split strips pivot tables and their caches by contract (packages/xlsx/src/contract.ts) and leaves external links unrewritten; describe what is preserved, what is removed, and what needs review, and say "the whole workbook" when you mean the mode rather than the guarantee',
  },
  {
    id: "every-tab-or-setting",
    // The trailing \b keeps "every table" out; the claim being linted is the
    // promise that all of a workbook's tabs or settings survive an edit.
    pattern: /every (?:tab|setting)s?\b/gi,
    advice:
      'name the structures actually kept ("sheets, formatting, and supported workbook structure") rather than promising all of them',
  },
  {
    id: "useful-worksheet",
    pattern: /useful worksheet/gi,
    advice:
      '"useful" hides the real rule: a worksheet is read when it is visible (or includeHiddenSheets is set), passes the sheets filter, and holds at least one non-blank row under a resolvable header',
  },
  {
    id: "stale-node-version",
    // Comparators are allowed so this phrase rule agrees with the engines
    // check below, which reads "Node.js >=24" as a version claim too.
    pattern: /Node(?:\.js)?\s*(?:>=|>|\^|~|=)?\s*v?24/gi,
    advice:
      'published packages declare engines.node ">=22.0.0"; a package README must state the supported floor, not the version the maintainers develop on',
    appliesTo: isPackageReadme,
  },
  {
    id: "counted-packages",
    pattern: /all (?:six|seven|eight) (?:tarballs|packages)/gi,
    advice:
      'a spelled-out count goes stale the moment a package is added; use count-free wording such as "a tarball for every published package"',
  },
  {
    id: "always-reflects",
    pattern: /always reflects/gi,
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

/**
 * A file flattened to one line, with the source line every character came
 * from. Whitespace is collapsed because a claim that wraps across a line break
 * is still a claim, and the parallel line index is what lets a wrapped match
 * still be reported - and allowlisted - at the line it starts on.
 */
interface FlattenedText {
  readonly lineOfCharacter: readonly number[];
  readonly text: string;
}

function flattenWhitespace(source: string): FlattenedText {
  const characters: string[] = [];
  const lineOfCharacter: number[] = [];
  let line = 1;
  let inWhitespaceRun = false;

  for (const character of source) {
    if (
      character === " " ||
      character === "\t" ||
      character === "\n" ||
      character === "\r"
    ) {
      if (!inWhitespaceRun) {
        characters.push(" ");
        lineOfCharacter.push(line);
        inWhitespaceRun = true;
      }
      if (character === "\n") {
        line += 1;
      }
      continue;
    }
    inWhitespaceRun = false;
    characters.push(character);
    lineOfCharacter.push(line);
  }

  return { lineOfCharacter, text: characters.join("") };
}

const problems: string[] = [];
const usedAllowlistEntries = new Set<string>();
// An exemption names one occurrence: this rule, matching this text, starting
// on this line of this file. A second occurrence of the same wording elsewhere
// in the file is a separate claim and is still reported.
const allowlistKey = (
  file: string,
  ruleId: string,
  line: number,
  phrase: string,
): string => `${file}::${ruleId}::${line}::${phrase.toLocaleLowerCase()}`;
const allowedKeys = new Set(
  ALLOWLIST.map((entry) =>
    allowlistKey(entry.file, entry.ruleId, entry.line, entry.phrase),
  ),
);

const scannedFiles = collectScannedFiles();

for (const absolutePath of scannedFiles) {
  const label = toRepoLabel(absolutePath);
  const { lineOfCharacter, text } = flattenWhitespace(
    readScannableText(absolutePath),
  );

  for (const rule of CLAIM_RULES) {
    if (rule.appliesTo && !rule.appliesTo(label)) {
      continue;
    }
    for (const match of text.matchAll(rule.pattern)) {
      const line = lineOfCharacter[match.index] ?? 1;
      const key = allowlistKey(label, rule.id, line, match[0]);
      if (allowedKeys.has(key)) {
        usedAllowlistEntries.add(key);
        continue;
      }
      problems.push(
        `${label}:${line} claims "${match[0]}" (rule ${rule.id})\n    ${rule.advice}`,
      );
    }
  }
}

for (const entry of ALLOWLIST) {
  const key = allowlistKey(entry.file, entry.ruleId, entry.line, entry.phrase);
  if (!usedAllowlistEntries.has(key)) {
    problems.push(
      `the claims allowlist entry for "${entry.phrase}" at ${entry.file}:${entry.line} (rule ${entry.ruleId}) matches nothing any more; move it to the line the claim is on now, or delete it\n    recorded reason: ${entry.reason}`,
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

// A README may name the version in prose ("Node.js 22 or later") or in the
// comparator syntax an engines field uses ("Node.js >=22"). Both forms are
// matched, so neither can slip past the check by not looking like a mention.
const nodeMentionPattern =
  /\bNode(?:\.js)?\s*(?<comparator>>=|>|\^|~|=)?\s*v?(?<version>\d+(?:\.\d+){0,2})/gi;
// "22+" is the one floor spelling that ends in a non-word character, so the
// word boundary belongs to the spelled-out alternatives only; requiring it
// after the plus rejected the conventional form outright.
const floorSuffixPattern =
  /^\s*(?:\+|(?:or (?:later|newer|above)|and later)\b)/i;

/** Pad a partial version to major.minor.patch so two spellings can be compared. */
function normalizeVersion(version: string): string {
  const [major = "0", minor = "0", patch = "0"] = version.split(".");
  return `${Number(major)}.${Number(minor)}.${Number(patch)}`;
}

/** The shortest spelling of a floor: 22.0.0 reads as 22, 22.13.0 as 22.13. */
function shortenVersion(version: string): string {
  return normalizeVersion(version).replace(/(?:\.0)+$/, "");
}
// The only comparator that states the same range an engines floor states.
// `>` excludes the declared minimum itself, and `^` and `~` cap the range at
// the next major, so each of them claims a support window the manifest does
// not, and each is reported with what it got wrong.
const INCLUSIVE_FLOOR_COMPARATOR = ">=";
const comparatorFaults: ReadonlyMap<string, string> = new Map([
  [">", "excludes the declared minimum itself"],
  ["^", "caps the range at the next major"],
  ["~", "caps the range at the next minor"],
  ["=", "pins one version"],
]);
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
    const declaredVersion = /(\d+(?:\.\d+){0,2})/.exec(declared ?? "")?.[1];
    const readmeLabel = toRepoLabel(readmePath);
    const manifestLabel = toRepoLabel(manifestPath);
    const readmeText = readTextFile(readmePath);

    const mentions = [...readmeText.matchAll(nodeMentionPattern)];
    if (mentions.length === 0) {
      continue;
    }
    if (declared === undefined || declaredVersion === undefined) {
      problems.push(
        `${readmeLabel} states a Node.js requirement, but ${manifestLabel} declares no engines.node to check it against`,
      );
      continue;
    }
    const floorLabel = shortenVersion(declaredVersion);

    for (const mention of mentions) {
      const line = readmeText.slice(0, mention.index).split("\n").length;
      const version = mention.groups?.version ?? "";
      // The whole version is compared, not just the major: a README promising
      // "22.13 or later" turns away the 22.0 to 22.12 readers that
      // engines.node ">=22.0.0" supports.
      if (normalizeVersion(version) !== normalizeVersion(declaredVersion)) {
        problems.push(
          `${readmeLabel}:${line} says "${mention[0]}", but ${manifestLabel} declares engines.node "${declared}"\n    state the version that field actually supports, or change engines.node if the support window really moved`,
        );
        continue;
      }
      const comparator = mention.groups?.comparator;
      if (comparator === INCLUSIVE_FLOOR_COMPARATOR) {
        continue;
      }
      const fault =
        comparator === undefined ? undefined : comparatorFaults.get(comparator);
      if (fault !== undefined) {
        problems.push(
          `${readmeLabel}:${line} says "${mention[0]}", but ${manifestLabel} declares engines.node "${declared}", and "${comparator}" ${fault}\n    write "Node.js ${floorLabel} or later", or ">=${floorLabel}", so the README states the same support window as the manifest`,
        );
        continue;
      }
      const remainder = readmeText.slice(mention.index + mention[0].length);
      if (!floorSuffixPattern.test(remainder)) {
        problems.push(
          `${readmeLabel}:${line} says "${mention[0]}" without saying it is a floor, but ${manifestLabel} declares engines.node "${declared}"\n    write "Node.js ${floorLabel} or later", or "Node.js ${floorLabel}+", so a reader on a newer Node is not turned away`,
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
