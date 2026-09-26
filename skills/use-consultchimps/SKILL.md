---
name: use-consultchimps
description:
  Run the ConsultChimps CLI for Excel, PowerPoint and PDF operations instead of
  writing a throwaway script. Covers invocation, running from a checkout, the
  --json envelope, refusals, warnings and errors without a code. Use before
  writing code that combines, splits, inspects, populates or unprotects Excel,
  PowerPoint or PDF files, or when the user mentions ConsultChimps.
license: Apache-2.0
metadata:
  cli-version: "0.12.0"
  repository: consultchimps/consultchimps
---

# Use the ConsultChimps CLI

This skill is the router: how to run the CLI and read what it returns. For Excel
recipes, load `chimps-xlsx` if it is installed; the commands it uses are also
listed in [references/cli-reference.md](references/cli-reference.md).

## Decision rule

Ten operations ship as tested CLI commands. For any of them, run the CLI. Do not
write a script that opens a workbook, walks a PDF, or edits OOXML by hand: the
CLI already handles hidden sheets, merged cells, cached formula values, and
refusing to overwrite its own input.

Write custom code only for work outside the list below, and say so, so the user
knows which part of the result the toolkit stands behind.

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

The `db` commands need native database bindings and are not covered here. In
0.12.0, `sheets consolidate` and `sheets merge` read `.xlsx` only; the other
`sheets` commands also take `.xlsm`.

"Combine these files" is ambiguous. Ask which shape the user wants: one table of
stacked rows (consolidate), or one workbook of separate tabs (merge).

## Invocation

```bash
npx consultchimps@0.12.0 <command>
```

No global install and no library import. The pinned version is the one this
skill documents: a flag missing from
[references/cli-reference.md](references/cli-reference.md) does not exist in it.

### Running from a checkout

Use the repository build when the user asks for it, or when a fix they need is
merged but not released.

```bash
npx pnpm@11.10.0 install --frozen-lockfile
npx pnpm@11.10.0 --filter "consultchimps..." build
node packages/cli/dist/index.js <command>
```

Take the pnpm version from `packageManager` in the root `package.json`. A build
prints the same `--version` as the release it follows, so compare the time of
`packages/cli/dist/index.js` with `git log -1` and rebuild when the build is
older. Behaviour described in this skill is that of 0.12.0; a checkout can
differ.

## Output contract

Human-readable results go to stdout, progress to stderr, and a failure sets a
nonzero exit code. Place `--json` before the command for one line of JSON:

```bash
npx consultchimps@0.12.0 --json pdf split report.pdf -o pages
```

Success prints `{"ok":true,"result":...}` with `artifacts`, `warnings` and
`metrics`. Failure prints `{"ok":false,"error":{"message":...,"code":...}}`.
Read the result, not the exit code alone, when you need counts or files created.

## Refusals, warnings and errors

**A refusal** carries a namespaced code such as `XLSX_SPLIT_COLUMN_NOT_FOUND`.
It stopped before writing anything. Fix the input or the option and run again;
do not retry unchanged and do not work around it with a script.

**A warning** means the operation completed and dropped or changed something on
the way, for example a column no mapping claimed. Report warnings to the user.

**An error with `"code": null`** is not a refusal: the CLI hit something it did
not anticipate, such as `Maximum call stack size exceeded`. It does not name the
input. Re-run on each input alone to find the file, tell the user it is a defect
in the tool, and check whether another operation reads that file (in 0.12.0,
`sheets inspect` and `sheets consolidate` use different readers).

The CLI never modifies an input file, and refuses to overwrite an existing
output unless `-f, --force` is given.

## Environment pitfalls

- **Windows long paths:** Git Bash and other MSYS shells cannot start `node`
  from a working directory longer than 260 characters ("path longer than allowed
  for a Win32 working directory"). Run from a short directory and pass long
  paths as arguments.
- **Locked files:** a workbook open in Excel, or held by OneDrive or SharePoint
  sync, fails to read or write with a lock error such as `EBUSY`, `EPERM` or
  `EACCES`. Ask the user to close it rather than copying around the lock.
- **Synced folders:** an output written into a OneDrive or SharePoint folder is
  visible to others as soon as it syncs. Write drafts elsewhere first.

## Vocabulary

| Term                  | Means                                                    |
| --------------------- | -------------------------------------------------------- |
| **Consolidate**       | stack rows from many worksheets into one table           |
| **Merge (workbooks)** | copy worksheets in as separate tabs, never stacking rows |
| **Split**             | one output file per distinct value, or per PDF page      |
| **Inspect**           | describe an input, producing no file                     |

## Related skills

Skills installed on their own do not ship each other's files, so these are
named, not linked.

- `chimps-xlsx`: consolidate, merge, split and inspect recipes, and column
  mapping
- `chimps-excel-design-skill`: authoring a new workbook deliverable
- `chimps-html-design-skill`: authoring a self-contained HTML deliverable
