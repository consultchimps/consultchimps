import { execFileSync } from "node:child_process";
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

// Typography check: no em dash (U+2014) and no en dash (U+2013) anywhere in the
// repository's tracked text. The repository writes short sentences that carry
// one idea each, and a dash is the punctuation that quietly defeats that: it
// glues two thoughts together, so the sentence keeps growing and the reader has
// to hold the first half open while the second arrives. Every occurrence this
// check found had a plainer reading waiting for it. A parenthetical becomes a
// comma pair, parentheses, or two sentences; a dash standing in for a colon
// becomes a colon; a dash joining a label to its explanation becomes a colon;
// a numeric range becomes "22 to 26".
//
// The point is the rewrite, not the substitution. Replacing an em dash with
// " - " or "--" keeps the sentence that needed splitting and adds an ASCII
// artefact to it, so neither is an accepted fix and neither is what this check
// is asking for.
//
// Scope is every tracked text file, which is wider than scripts/check-claims.ts
// deliberately: a claim is only wrong where a user reads it, while a dash is a
// house-style violation wherever it is written, and a CHANGELOG entry is
// rendered on the releases page, a code comment is read by the next
// contributor, and a workflow comment by whoever debugs the release. Binary
// files, the lockfile, and everything git does not track are out of scope.
//
// This is a character scan and it is meant to stay one. It does not judge
// prose, count sentence length, or police any other punctuation, because a rule
// that cannot be checked by looking at one character is a rule a reviewer
// should be applying instead. Widening it into a general prose linter would
// trade a check that is always right for one that is usually annoying.
//
// A genuine exception goes in the allowlist below with the reason it is
// deliberate, keyed on the file, the line, and the character. There are none
// today: every dash in the repository's history was a sentence that read better
// without it.

/** A dash this check bans, and the plainer punctuation that replaces it. */
interface BannedDash {
  /** The character itself, so a message can name it exactly. */
  readonly character: string;
  /** How the character is spelled in a failure message. */
  readonly name: string;
  /** What to write instead. */
  readonly advice: string;
}

interface AllowlistEntry {
  /** Repository-relative path, forward slashes. */
  readonly file: string;
  /**
   * One-based line the dash sits on. Together with the character it identifies
   * one occurrence, so an exemption can never widen into a blanket pardon: a
   * second dash further down the same file is still reported. A line number
   * moves, but moving it can only make the dash reappear as a failure and the
   * entry report itself stale, never let an unexamined dash pass quietly.
   */
  readonly line: number;
  /** The banned character being excused, as the character itself. */
  readonly character: string;
  readonly reason: string;
}

const workspaceRoot = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
);

// The characters are written as escapes so this file, which the check scans
// like any other, does not report itself.
const BANNED_DASHES: readonly BannedDash[] = [
  {
    advice:
      'rewrite the sentence: a parenthetical takes a comma pair, parentheses, or a full stop; a dash standing in for a colon takes a colon; a label and its explanation take a colon. Never substitute " - " or "--" in prose',
    character: "\u2014",
    name: "em dash (U+2014)",
  },
  {
    advice:
      'rewrite the sentence, or write a numeric range as "22 to 26"; a code-like range takes a plain hyphen',
    character: "\u2013",
    name: "en dash (U+2013)",
  },
];

// Deliberate occurrences, each with the reason it is not a style violation.
// Keep this list short: needing many entries means the house style has changed,
// not that a file deserves an exemption.
const ALLOWLIST: readonly AllowlistEntry[] = [];

// Text this repository writes. Anything else tracked in git is data, an image,
// or a lockfile, where a dash is not the repository's prose to fix.
const TEXT_EXTENSIONS: readonly string[] = [
  ".css",
  ".json",
  ".md",
  ".mdx",
  ".ts",
  ".tsx",
  ".txt",
  ".yaml",
  ".yml",
];

/**
 * Every tracked text file, from git rather than a directory walk: git already
 * knows what is ignored, generated, or vendored, so the check cannot drift out
 * of step with .gitignore, and a file the repository does not track is not the
 * repository's prose. Names are NUL-separated so a path with a space or a
 * non-ASCII character survives.
 */
function listTrackedTextFiles(): string[] {
  const output = execFileSync("git", ["ls-files", "-z"], {
    cwd: workspaceRoot,
    encoding: "utf8",
    maxBuffer: 64 * 1024 * 1024,
  });
  return output
    .split("\0")
    .filter((label) => label !== "")
    .filter((label) =>
      TEXT_EXTENSIONS.some((extension) => label.endsWith(extension)),
    )
    .filter((label) => existsSync(path.join(workspaceRoot, label)))
    .sort();
}

/** Enough of the offending line to see the sentence the dash sits in. */
function excerpt(line: string): string {
  const trimmed = line.trim();
  return trimmed.length > 120 ? `${trimmed.slice(0, 117)}...` : trimmed;
}

const problems: string[] = [];
const usedAllowlistEntries = new Set<string>();
// An exemption names one occurrence: this character, on this line of this file.
// Another dash elsewhere in the file is a separate violation and is reported.
const allowlistKey = (file: string, line: number, character: string): string =>
  `${file}::${line}::${character}`;
const allowedKeys = new Set(
  ALLOWLIST.map((entry) =>
    allowlistKey(entry.file, entry.line, entry.character),
  ),
);

const scannedFiles = listTrackedTextFiles();

for (const label of scannedFiles) {
  const absolutePath = path.join(workspaceRoot, ...label.split("/"));
  const text = readFileSync(absolutePath, "utf8");
  if (!BANNED_DASHES.some((dash) => text.includes(dash.character))) {
    continue;
  }

  const lines = text.replace(/\r\n/g, "\n").split("\n");
  for (const [index, line] of lines.entries()) {
    for (const dash of BANNED_DASHES) {
      if (!line.includes(dash.character)) {
        continue;
      }
      const lineNumber = index + 1;
      const key = allowlistKey(label, lineNumber, dash.character);
      if (allowedKeys.has(key)) {
        usedAllowlistEntries.add(key);
        continue;
      }
      problems.push(
        `${label}:${lineNumber} contains an ${dash.name}\n    ${excerpt(line)}\n    ${dash.advice}`,
      );
    }
  }
}

for (const entry of ALLOWLIST) {
  const key = allowlistKey(entry.file, entry.line, entry.character);
  if (!usedAllowlistEntries.has(key)) {
    problems.push(
      `the typography allowlist entry at ${entry.file}:${entry.line} matches no dash any more; move it to the line the dash is on now, or delete it\n    recorded reason: ${entry.reason}`,
    );
  }
}

if (problems.length > 0) {
  throw new Error(
    `Tracked text contains dashes the house style does not use:\n${problems
      .map((problem) => `- ${problem}`)
      .join(
        "\n",
      )}\nRewrite each sentence so it reads plainly without the dash, or add an allowlist entry in scripts/check-typography.ts with the reason the character is deliberate.`,
  );
}

process.stdout.write(
  `Checked ${scannedFiles.length} tracked text files for em and en dashes, with ${ALLOWLIST.length} allowlisted occurrences.\n`,
);
