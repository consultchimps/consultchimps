---
"consultchimps": patch
"@consultchimps/core": patch
"@consultchimps/files": patch
"@consultchimps/messages": patch
"@consultchimps/pdf": patch
"@consultchimps/pptx": patch
"@consultchimps/tabular": patch
"@consultchimps/xlsx": patch
---

Declare support for Node.js 22 and later (`engines.node: ">=22.0.0"` instead of
`24.x`), so the toolkit installs and runs in environments that ship the previous
LTS line. CI now validates the runtime on Node 22.16, the latest 22, and 26
alongside the full Node 24 verification. The CLI additionally ships a standalone
`consultchimps.mjs` bundle on each GitHub release: one file that runs with
`node consultchimps.mjs` and needs no npm access at all.
