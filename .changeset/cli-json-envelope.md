---
"consultchimps": minor
---

Give `--json` a single-line result envelope on stdout. A successful command now
prints `{"ok":true,"result":...}` and a failing command prints
`{"ok":false,"error":{"message":...,"code":...}}` while keeping its nonzero exit
code, so an automation can branch on `ok` and read the stable error code without
parsing human text. Usage errors such as an unknown option or a missing required
option are reported through the same envelope with the code `CLI_USAGE`, instead
of leaving stdout empty. `--help` and `--version` are unchanged.

This changes the shape of `--json` output. Commands previously printed the
operation result pretty-printed and unwrapped, and a failure wrote a prose line
to stderr with nothing machine-readable on stdout. Consumers that read `--json`
should now read the value under `result`.
