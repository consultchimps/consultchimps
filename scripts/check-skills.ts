import { existsSync, readFileSync, readdirSync, statSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

// Skills check: every directory under skills/ is a skill a registry can list
// and install on its own. skills.sh installs with `npx skills add --skill
// <name>`, which copies one skill directory and nothing beside it, and agents
// that load a skill trust its frontmatter to follow the Agent Skills
// specification (https://agentskills.io/specification). A skill that breaks
// either still looks fine in this repository, because its neighbours are
// there; it only fails after install, on someone else's machine.
//
// The frontmatter reader handles the subset of YAML these skills use: scalar
// keys, plain scalars folded over indented lines, and a one level `metadata`
// map. Anything else stops the check rather than being skipped, so a skill
// cannot pass by using syntax this reader does not understand.
//
// A name ending in "-skill" marks a craft skill meant for upload into ChatGPT
// and Claude chats as well as for the registry. Claude.ai rejects an upload
// whose description is longer than 200 characters, so those skills are held
// to that limit rather than the specification's 1,024.
//
// Each skill carries its own LICENSE.txt, a copy of the repository's LICENSE,
// and declares that license in its frontmatter. A skill installed alone or
// zipped for a chat leaves the repository behind, so the terms have to travel
// inside the skill directory, as they do in Anthropic's and OpenAI's skill
// repositories.
//
// Plain (unquoted) scalars are held to YAML's own rules as well, because a real
// YAML parser is what reads them after install. A description containing ": "
// looks like a nested mapping to YAML, and the `skills` CLI skips the whole
// skill with a parse error, so skills.sh never lists it. Two skills shipped
// that way before this rule existed.

const workspaceRoot = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
);
const skillsRoot = path.join(workspaceRoot, "skills");

const normalized = (text: string): string => text.replace(/\r\n/g, "\n");
const repositoryLicenseText = normalized(
  readFileSync(path.join(workspaceRoot, "LICENSE"), "utf8"),
);
const repositoryLicense: unknown = (
  JSON.parse(
    readFileSync(path.join(workspaceRoot, "package.json"), "utf8"),
  ) as { license?: unknown }
).license;
if (typeof repositoryLicense !== "string") {
  throw new Error("package.json declares no license string.");
}

const NAME_PATTERN = /^[a-z0-9]+(?:-[a-z0-9]+)*$/;
const NAME_MAX = 64;
const DESCRIPTION_MAX = 1024;
const CHAT_UPLOAD_SUFFIX = "-skill";
const CHAT_UPLOAD_DESCRIPTION_MAX = 200;
const COMPATIBILITY_MAX = 500;
const SKILL_MD_MAX_LINES = 500;
const KNOWN_KEYS = new Set([
  "name",
  "description",
  "license",
  "compatibility",
  "metadata",
  "allowed-tools",
]);

/**
 * Why a plain scalar would not parse as one string, or null when it would.
 * Quoted values are exempt: inside quotes ": " and " #" are ordinary text.
 */
function plainScalarProblem(value: string): string | null {
  if (/^["']/.test(value)) {
    return null;
  }
  if (/^[-?:,[\]{}#&*!|>'"%@`]/.test(value)) {
    return `it starts with "${value[0] ?? ""}", which YAML reads as syntax`;
  }
  if (value.includes(": ") || value.endsWith(":")) {
    return 'it contains ": ", which YAML reads as a nested mapping';
  }
  if (value.includes(" #")) {
    return 'it contains " #", which YAML reads as the start of a comment';
  }
  return null;
}

interface Frontmatter {
  readonly fields: Map<string, string>;
  readonly metadata: Map<string, string>;
}

function unquote(value: string): string {
  const trimmed = value.trim();
  if (
    trimmed.length >= 2 &&
    (trimmed.startsWith('"') || trimmed.startsWith("'")) &&
    trimmed.endsWith(trimmed[0] ?? "")
  ) {
    return trimmed.slice(1, -1);
  }
  return trimmed;
}

function parseFrontmatter(text: string, label: string): Frontmatter {
  const lines = text.replace(/\r\n/g, "\n").split("\n");
  if (lines[0] !== "---") {
    throw new Error(
      `${label}: SKILL.md must start with a --- frontmatter line.`,
    );
  }
  const end = lines.indexOf("---", 1);
  if (end === -1) {
    throw new Error(`${label}: the frontmatter has no closing --- line.`);
  }

  const fields = new Map<string, string>();
  const plain = new Set<string>();
  const metadata = new Map<string, string>();
  let current: string | null = null;
  let inMetadata = false;

  for (const line of lines.slice(1, end)) {
    if (line.trim() === "") {
      continue;
    }
    const topLevel = /^([A-Za-z][\w-]*):(.*)$/.exec(line);
    if (topLevel?.[1] !== undefined) {
      const key = topLevel[1];
      const value = (topLevel[2] ?? "").trim();
      inMetadata = key === "metadata";
      current = inMetadata ? null : key;
      if (!inMetadata) {
        fields.set(key, unquote(value));
        if (!/^["']/.test(value)) {
          plain.add(key);
        }
      } else if (value !== "") {
        throw new Error(`${label}: metadata must be a map, one key per line.`);
      }
      continue;
    }
    if (inMetadata) {
      const entry = /^ {2}([\w-]+):\s*(.*)$/.exec(line);
      if (entry?.[1] === undefined) {
        throw new Error(`${label}: cannot read metadata line "${line}".`);
      }
      const raw = (entry[2] ?? "").trim();
      if (raw === "" || /^[[{|>]/.test(raw)) {
        throw new Error(
          `${label}: metadata.${entry[1]} must be a single string value.`,
        );
      }
      if (
        !/^["']/.test(raw) &&
        /^(?:-?\d+(?:\.\d+)?|true|false|null)$/.test(raw)
      ) {
        throw new Error(
          `${label}: metadata.${entry[1]} reads as a ${raw === "true" || raw === "false" ? "boolean" : "number"}; quote it, because the specification maps metadata keys to strings.`,
        );
      }
      const problem = plainScalarProblem(raw);
      if (problem !== null) {
        throw new Error(
          `${label}: metadata.${entry[1]} does not parse as a string: ${problem}. Quote it.`,
        );
      }
      metadata.set(entry[1], unquote(raw));
      continue;
    }
    if (current !== null && /^ {2}\S/.test(line)) {
      const previous = fields.get(current) ?? "";
      fields.set(current, `${previous} ${line.trim()}`.trim());
      continue;
    }
    throw new Error(`${label}: cannot read frontmatter line "${line}".`);
  }

  for (const key of plain) {
    const problem = plainScalarProblem(fields.get(key) ?? "");
    if (problem !== null) {
      throw new Error(
        `${label}: ${key} does not parse as one string: ${problem}. Reword it, or quote the whole value.`,
      );
    }
  }

  return { fields, metadata };
}

/** Relative link targets in a Markdown file, without anchors. */
function linkTargets(markdown: string): string[] {
  const targets: string[] = [];
  for (const match of markdown.matchAll(/\]\(([^)\s]+)(?:\s+"[^"]*")?\)/g)) {
    const target = match[1] ?? "";
    if (/^(?:[a-z]+:|#|\/)/i.test(target)) {
      continue;
    }
    targets.push(target.split("#")[0] ?? "");
  }
  return targets.filter((target) => target !== "");
}

function markdownFiles(directory: string): string[] {
  return readdirSync(directory).flatMap((entry) => {
    const full = path.join(directory, entry);
    if (statSync(full).isDirectory()) {
      return markdownFiles(full);
    }
    return entry.endsWith(".md") ? [full] : [];
  });
}

const problems: string[] = [];
const skills = readdirSync(skillsRoot).filter((entry) =>
  statSync(path.join(skillsRoot, entry)).isDirectory(),
);

for (const skill of skills) {
  const skillDirectory = path.join(skillsRoot, skill);
  const skillFile = path.join(skillDirectory, "SKILL.md");
  const label = `skills/${skill}`;
  if (!existsSync(skillFile)) {
    problems.push(`${label}: no SKILL.md.`);
    continue;
  }
  const text = readFileSync(skillFile, "utf8");

  let frontmatter: Frontmatter;
  try {
    frontmatter = parseFrontmatter(text, label);
  } catch (error) {
    problems.push((error as Error).message);
    continue;
  }
  const { fields } = frontmatter;

  for (const key of fields.keys()) {
    if (!KNOWN_KEYS.has(key)) {
      problems.push(
        `${label}: "${key}" is not an Agent Skills frontmatter field; put it under metadata.`,
      );
    }
  }

  const name = fields.get("name") ?? "";
  if (!NAME_PATTERN.test(name) || name.length > NAME_MAX) {
    problems.push(
      `${label}: name "${name}" must be 1 to ${String(NAME_MAX)} lowercase letters, digits and single hyphens.`,
    );
  }
  if (name !== skill) {
    problems.push(`${label}: name "${name}" must match the directory name.`);
  }

  const description = fields.get("description") ?? "";
  if (description.length === 0 || description.length > DESCRIPTION_MAX) {
    problems.push(
      `${label}: description is ${String(description.length)} characters; it must be 1 to ${String(DESCRIPTION_MAX)}.`,
    );
  }

  if (
    name.endsWith(CHAT_UPLOAD_SUFFIX) &&
    description.length > CHAT_UPLOAD_DESCRIPTION_MAX
  ) {
    problems.push(
      `${label}: description is ${String(description.length)} characters; a "${CHAT_UPLOAD_SUFFIX}" skill is uploaded to Claude.ai, which accepts at most ${String(CHAT_UPLOAD_DESCRIPTION_MAX)}.`,
    );
  }

  const license = fields.get("license");
  if (license !== repositoryLicense) {
    problems.push(
      `${label}: license is "${license ?? ""}"; it must be "${repositoryLicense}", the repository's license.`,
    );
  }
  const licenseFile = path.join(skillDirectory, "LICENSE.txt");
  if (!existsSync(licenseFile)) {
    problems.push(
      `${label}: no LICENSE.txt. Copy the repository's LICENSE into the skill directory so the terms travel with it.`,
    );
  } else if (
    normalized(readFileSync(licenseFile, "utf8")) !== repositoryLicenseText
  ) {
    problems.push(
      `${label}: LICENSE.txt differs from the repository's LICENSE. Copy it again.`,
    );
  }

  const compatibility = fields.get("compatibility");
  if (
    compatibility !== undefined &&
    (compatibility.length === 0 || compatibility.length > COMPATIBILITY_MAX)
  ) {
    problems.push(
      `${label}: compatibility must be 1 to ${String(COMPATIBILITY_MAX)} characters.`,
    );
  }

  const lineCount = text.split("\n").length;
  if (lineCount > SKILL_MD_MAX_LINES) {
    problems.push(
      `${label}: SKILL.md is ${String(lineCount)} lines; keep it under ${String(SKILL_MD_MAX_LINES)} and move detail into references/.`,
    );
  }

  // A link must resolve inside this skill's own directory: after a single
  // skill install, nothing outside it exists.
  for (const file of markdownFiles(skillDirectory)) {
    const fileLabel = path
      .relative(workspaceRoot, file)
      .split(path.sep)
      .join("/");
    for (const target of linkTargets(readFileSync(file, "utf8"))) {
      const resolved = path.resolve(path.dirname(file), target);
      const inside = path.relative(skillDirectory, resolved);
      if (inside.startsWith("..") || path.isAbsolute(inside)) {
        problems.push(
          `${fileLabel}: link "${target}" leaves the skill directory, so it breaks when the skill is installed alone. Name the other skill instead.`,
        );
      } else if (!existsSync(resolved)) {
        problems.push(`${fileLabel}: link "${target}" points at nothing.`);
      }
    }
  }
}

if (problems.length > 0) {
  process.stderr.write(
    `Skills check failed:\n${problems.map((problem) => `  - ${problem}`).join("\n")}\n`,
  );
  process.exit(1);
}

process.stdout.write(
  `Verified ${String(skills.length)} skills against the Agent Skills specification and single skill install.\n`,
);
