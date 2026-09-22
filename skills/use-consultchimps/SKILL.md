---
name: use-consultchimps
description:
  Run the ConsultChimps CLI for Excel, PowerPoint and PDF operations instead of
  writing a throwaway script. Covers invocation, the --json envelope, exit
  codes, reading refusals and warnings, and the toolkit vocabulary. Use before
  writing any code that combines, splits, inspects, populates or unprotects
  .xlsx, .xlsm, .pptx or .pdf files, and whenever the user mentions
  ConsultChimps or consultchimps.
license: MIT
metadata:
  cli-version: "0.12.0"
  repository: consultchimps/consultchimps
---

# Use the ConsultChimps CLI

## Decision rule

Ten operations ship as tested CLI commands. For any of them, run the CLI. Do not
write a script that opens a workbook, walks a PDF, or edits OOXML by hand: the
CLI already handles the cases a script written in one sitting will miss,
including hidden sheets, merged cells, cached formula values, and refusing to
overwrite its own input.

Write custom code only for work outside the list below. When you fall back to
custom code, say so and say why, so the user knows which part of the result the
toolkit stands behind.

## Invocation

```bash
npx consultchimps@0.12.0 <command>
```

No global install and no library import. The pinned version is the one this
skill documents; `references/cli-reference.md` is generated from that build, so
a flag it does not list does not exist in that version.

## The registered operations

| Task                                        | Command                 |
| ------------------------------------------- | ----------------------- |
| Stack rows from many worksheets into one    | `sheets consolidate`    |
| Copy worksheets into one workbook as tabs   | `sheets merge`          |
| One workbook per distinct value of a column | `sheets split`          |
| Describe a workbook without writing a file  | `sheets inspect`        |
| Remove worksheet protection                 | `sheets unprotect`      |
| Fill a PowerPoint template per data row     | `pptx populate`         |
| Describe a PowerPoint template              | `pptx inspect-template` |
| One file per page of a PDF                  | `pdf split`             |
| Combine PDFs into one file                  | `pdf merge`             |
| Prepare a reviewed database import batch    | `db import prepare`     |

The `db` commands need native database bindings and are not covered by recipes
here. The other nine run anywhere Node runs.

## Output contract

Human-readable results go to stdout, progress and diagnostics to stderr, and a
failure sets a nonzero exit code.

Place `--json` before the command to replace the explanation with one line of
JSON on stdout:

```bash
npx consultchimps@0.12.0 --json pdf split report.pdf -o pages
```

Success prints `{"ok":true,"result":...}`. Failure prints
`{"ok":false,"error":{"message":...,"code":...}}` and still exits nonzero.
Nothing else reaches stdout, so the output pipes straight into a parser. Read
the result rather than the exit code alone when you need counts or the list of
files created.

## Refusals and warnings

A refusal is deliberate and carries a stable, namespaced code such as
`XLSX_SPLIT_COLUMN_NOT_FOUND`. It means the operation stopped before writing
anything. Fix the input or the option and run again; do not retry unchanged and
do not work around it with a script.

A warning means the operation completed and something was dropped or changed on
the way, for example a column no mapping claimed, or a structure a merge cannot
carry. Warnings are part of the result. Report them to the user rather than
discarding them.

The CLI never modifies an input file, and refuses to overwrite an existing
output unless `-f, --force` is given.

## Vocabulary

These four are confused constantly. The toolkit uses them with exactly these
meanings, and so should you when talking to the user.

| Term                  | Means                                                    |
| --------------------- | -------------------------------------------------------- |
| **Consolidate**       | stack rows from many worksheets into one table           |
| **Merge (workbooks)** | copy worksheets in as separate tabs, never stacking rows |
| **Split**             | one output file per distinct value, or per PDF page      |
| **Inspect**           | describe an input, producing no file                     |

"Combine these files" is ambiguous. Ask which shape the user wants: one table of
stacked rows, or one workbook of separate tabs.

## Going deeper

- Excel work in depth, including choosing between operations and mapping
  mismatched headers: the `chimps-xlsx` skill
- Authoring a new workbook as a deliverable rather than transforming one: the
  `chimps-excel-design-skill` skill
- Every command and option of the pinned version:
  [the generated CLI reference](references/cli-reference.md)
