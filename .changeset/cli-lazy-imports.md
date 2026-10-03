---
"consultchimps": patch
---

Start in a fraction of a second. The CLI loaded every package before reading its
command, including the two native database engines, so `--version`, `--help`,
and every command took several seconds to begin; on a busy Windows machine it
measured 13 to 27 seconds. Each command now loads only the packages it uses,
when it runs, and `--version` takes about 0.2 seconds.
