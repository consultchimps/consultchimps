---
"@consultchimps/messages": minor
"consultchimps": patch
---

Let every interface reuse the plain-language explanations in its own words.
`formatHumanResult` and `formatHumanError` now accept an optional
`{ vocabulary }` option typed as the new exported `MessageVocabulary`, which
holds the interface-specific phrases: how to retry with overwriting enabled,
where to find the reference or examples, how to inspect a PowerPoint template
first, how to point at the created files, and the word for a unit of work.

Two vocabularies ship with the package. `GENERIC_VOCABULARY` is the new default
and never names a flag, an executable, or a terminal, so a desktop or browser
interface can show the guidance unchanged. `CLI_VOCABULARY` reproduces the
command-line wording, including `--force`, `--help`, and
`consultchimps pptx inspect-template`.

Both functions keep their existing signatures, so current callers still compile.
Library callers that do not pass a vocabulary now receive the neutral wording
instead of command-line instructions. The `consultchimps` CLI passes
`CLI_VOCABULARY`, so its output is byte-for-byte unchanged.
