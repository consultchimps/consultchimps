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

The two tool skills drive the published CLI. The two craft skills are standards
for consulting work and need no ConsultChimps install, though they use it when
it is there.

## Install

From this repository, as a Claude Code plugin:

```bash
/plugin install consultchimps@consultchimps/consultchimps
```

Or with the registry CLI, which reads the same directory:

```bash
npx skills add consultchimps/consultchimps
```

## The generated CLI reference

`use-consultchimps/references/cli-reference.md` is generated from the built CLI
and must never be hand-edited. A skill that names a flag the CLI does not have
sends an agent into an error it cannot diagnose, so `pnpm docs:check`
regenerates the file and fails when the committed copy differs.

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
- claim nothing the tools cannot do today
- no em dashes or en dashes anywhere; `pnpm docs:check` fails on one
