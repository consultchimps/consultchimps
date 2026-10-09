# CLI golden output

Each `.txt` file is the exact human-readable output of one command run by
`../golden.test.ts`: the command line, exit code, stdout, and stderr. `help/`
holds every `--help` page, found by walking the help tree; `run/` holds runs on
small generated fixtures.

A changed golden is a changed user-facing message, so review its diff like any
other wording change.

## Updating

Build the CLI, then rewrite the goldens and review the diff:

```bash
pnpm --filter "consultchimps..." build
pnpm vitest run packages/cli/test/golden.test.ts -u
```

A new command needs a help golden, and a test fails until it has one. Run with
`-u` twice when adding one: the first pass writes the file, the second confirms
the set.

## What is normalised

Only fields that differ between machines:

- The case's temporary folder becomes `<tmp>`.
- On Windows, `\` becomes `/` inside those `<tmp>` paths only.
- CRLF becomes LF.
- Terminal control sequences are removed.

The environment is pinned instead of normalised: `NO_COLOR=1`, `COLUMNS=100`,
`TZ=UTC`, `LANG` and `LC_ALL` at `en_US.UTF-8` (not `C`, which ICU on Linux
reads as a POSIX locale that sorts unlike Windows), pipes rather than a
terminal, and `CONSULTCHIMPS_LOG=off` so no run record is written. No covered
output prints a duration, date, run id, or version; a case that does must add
its normaliser here and in the test.
