# Agent skills

Skills in the [Agent Skills](https://agentskills.io/specification) format: one
directory per skill, each with a `SKILL.md` carrying `name` and `description`
frontmatter, and detail in `references/` and `assets/` beside it. They work in
Claude Code, Codex and ChatGPT alike; nothing here is specific to one client.

| Skill                       | Kind  | Covers                                                              |
| --------------------------- | ----- | ------------------------------------------------------------------- |
| `use-consultchimps`         | tool  | invocation, the `--json` envelope, exit codes, refusals, vocabulary |
| `chimps-xlsx`               | tool  | transforming existing workbooks: consolidate, merge, split, inspect |
| `chimps-excel-design-skill` | craft | authoring a workbook deliverable: formulas, layout, handover        |
| `chimps-html-design-skill`  | craft | authoring a self-contained HTML deliverable: charts, theme, RTL     |

The two tool skills drive the published CLI, so they belong in an agent with a
terminal: Claude Code, Codex and similar. The two craft skills are standards for
consulting work and need no ConsultChimps install, though they use it when it is
there. They are also meant for upload into ChatGPT and Claude chats, and the
`-skill` ending on their names marks that.

## Install

With the Agent Skills CLI, which is how [skills.sh](https://skills.sh) lists
them:

```bash
npx skills add consultchimps/consultchimps                     # choose interactively
npx skills add consultchimps/consultchimps --skill chimps-xlsx # one skill
npx skills add consultchimps/consultchimps --list              # see what is here
```

Or from this repository, as a Claude Code plugin:

```bash
/plugin install consultchimps@consultchimps/consultchimps
```

### Uploading a craft skill to ChatGPT or Claude

Zip the skill's directory so the directory itself is the root of the archive,
for example `chimps-html-design-skill.zip` holding
`chimps-html-design-skill/SKILL.md` and its `assets/` and `references/`. Then
upload it under Skills in the ChatGPT or Claude settings.

Claude.ai rejects a description longer than 200 characters, so a `-skill`
description stays within 200; `scripts/check-skills.ts` enforces it. ChatGPT has
no fixed limit, but the list of every installed skill shares a small part of the
context, so a short description that leads with its trigger words is what gets
the skill chosen there too.

### Pointing a project at the skills

An agent reads a skill's description only when deciding whether to load it. In a
project that relies on these skills, a line in its `AGENTS.md` or `CLAUDE.md`
such as "Use the ConsultChimps skills for Excel, PowerPoint and PDF work" makes
the agent look for them.

## Publishing on skills.sh

skills.sh indexes a public repository from installs made with `npx skills add`;
there is no submission step. What it needs from this directory:

- Each skill stands alone. `npx skills add --skill <name>` copies one skill
  directory and nothing else, so a skill links only to files inside its own
  directory and names other skills rather than linking to them.
- Frontmatter follows the
  [Agent Skills specification](https://agentskills.io/specification): `name` of
  lowercase letters, digits and single hyphens, matching the directory;
  `description` of at most 1,024 characters; `metadata` values as strings.
- `SKILL.md` stays under 500 lines, with detail in `references/`, linked one
  level deep.

`pnpm docs:check` runs `scripts/check-skills.ts`, which fails on a broken rule
above.

## The generated CLI reference

`references/cli-reference.md` in `use-consultchimps` and in `chimps-xlsx` is
generated from the built CLI and must never be hand-edited. Each tool skill
carries its own copy so that it works when installed alone. A skill that names a
flag the CLI does not have sends an agent into an error it cannot diagnose, so
`pnpm docs:check` regenerates the file and fails when the committed copy
differs.

```bash
pnpm build          # the reference is read from packages/cli/dist
pnpm skills:reference
```

Both tool skills pin the CLI version they document in frontmatter. The pin is a
minimum: a later CLI keeps working, and the drift check catches the day its
command surface stops matching.

## House rules for editing a skill

- `name` must match the directory name exactly
- `description` says what the skill does and when to use it, in one paragraph,
  because it is all an agent sees until the skill fires
- keep `SKILL.md` short and push detail into `references/`, which is loaded only
  when needed
- `description` leads with what the skill does and the words a user would say;
  rules and refusals go in the body. A `-skill` description stays within 200
  characters for Claude.ai uploads; a tool skill may run to about 60 words
- a name ends in `-skill` only when the skill is meant for chat upload
- a claim about how the CLI behaves names the version it was checked against,
  because the drift check covers flags, not behaviour; re-check those claims
  when a release changes a reader
- link only inside the skill's own directory
- claim nothing the tools cannot do today
- no em dashes or en dashes anywhere; `pnpm docs:check` fails on one
