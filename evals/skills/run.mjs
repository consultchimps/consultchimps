// Runs each skills eval case through a fresh headless Claude Code agent, one at
// a time, with this repository loaded as a plugin, and writes a trace and a
// summary per case. A fallback for machines where `claude plugin eval` cannot
// grant Bash; the cases are in its format, so either runner reads them.
//
//   node evals/skills/run.mjs <out-dir> [case-name-substring ...]
//
// Needs claude, bash, node, python with openpyxl, and network for npx.
import { spawnSync } from "node:child_process";
import { mkdirSync, readFileSync, readdirSync, writeFileSync } from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const repo = path.resolve(here, "../..");
const [outDir, ...filters] = process.argv.slice(2);
if (!outDir) {
  process.stderr.write("usage: node evals/skills/run.mjs <out-dir> [case]\n");
  process.exit(2);
}

const cases = readdirSync(here, { withFileTypes: true })
  .filter((entry) => entry.isDirectory() && /^\d\d-/.test(entry.name))
  .map((entry) => entry.name)
  .filter(
    (name) => filters.length === 0 || filters.some((f) => name.includes(f)),
  );

function frontmatter(text) {
  const match = /^---\n([\s\S]*?)\n---\n([\s\S]*)$/.exec(
    text.replace(/\r\n/g, "\n"),
  );
  const fields = Object.fromEntries(
    match[1]
      .split("\n")
      .map((line) => line.split(/:\s*/, 2))
      .filter((pair) => pair.length === 2),
  );
  return { fields, body: match[2].trim() };
}

function summarize(trace) {
  const skills = [];
  const commands = [];
  let tools = 0;
  let result = {};
  for (const line of trace.split("\n")) {
    if (!line.trim()) continue;
    let event;
    try {
      event = JSON.parse(line);
    } catch {
      continue;
    }
    if (event.type === "result") result = event;
    if (event.type !== "assistant") continue;
    for (const block of event.message?.content ?? []) {
      if (block.type !== "tool_use") continue;
      tools += 1;
      if (block.name === "Skill") skills.push(block.input.skill);
      else if (block.name === "Bash" || block.name === "PowerShell")
        commands.push(block.input.command);
      else
        commands.push(
          `[${block.name}] ${block.input.file_path ?? block.input.pattern ?? ""}`,
        );
    }
  }
  return { skills, commands, tools, result };
}

for (const name of cases) {
  const caseDir = path.join(here, name);
  const { fields, body } = frontmatter(
    readFileSync(path.join(caseDir, "prompt.md"), "utf8"),
  );
  const work = path.resolve(outDir, name, "work");
  mkdirSync(work, { recursive: true });
  spawnSync("bash", [path.join(caseDir, "scaffold.sh").replaceAll("\\", "/")], {
    cwd: work,
    stdio: "inherit",
  });
  const started = Date.now();
  const run = spawnSync(
    "claude",
    [
      "-p",
      body,
      "--model",
      process.env.EVAL_MODEL ?? "opus",
      "--plugin-dir",
      repo,
      "--setting-sources",
      "project",
      "--strict-mcp-config",
      "--dangerously-skip-permissions",
      "--no-session-persistence",
      "--max-turns",
      fields.max_turns ?? "40",
      "--output-format",
      "stream-json",
      "--verbose",
    ],
    {
      cwd: work,
      encoding: "utf8",
      maxBuffer: 256 * 1024 * 1024,
      timeout: 30 * 60 * 1000,
    },
  );
  const trace = run.stdout ?? "";
  writeFileSync(path.resolve(outDir, name, "trace.jsonl"), trace);
  const { skills, commands, tools, result } = summarize(trace);
  const minutes = ((Date.now() - started) / 60000).toFixed(1);
  const summary = [
    `# ${name}`,
    `Prompt: ${body}`,
    `Skills: ${skills.join(", ") || "none"}`,
    `Tool calls: ${tools}; turns: ${result.num_turns ?? "?"}; minutes: ${minutes}; cost: ${result.total_cost_usd ?? "?"}`,
    "## Commands",
    ...commands.map((c) => `- ${c.replace(/\s+/g, " ").slice(0, 400)}`),
    "## Final message",
    result.result ?? `(no result; exit ${run.status}) ${run.stderr ?? ""}`,
  ].join("\n");
  writeFileSync(path.resolve(outDir, name, "summary.md"), `${summary}\n`);
  process.stderr.write(
    `${name}: skills=[${skills.join(", ")}] tools=${tools} ${minutes}m\n`,
  );
}
