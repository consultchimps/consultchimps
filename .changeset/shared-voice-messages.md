---
"@consultchimps/messages": minor
"consultchimps": patch
---

Extract the CLI's plain-language result and error rendering into the new
`@consultchimps/messages` package so desktop and web interfaces can reuse the
same explanations. The CLI output is unchanged apart from a new recovery
explanation for cancelled operations.
