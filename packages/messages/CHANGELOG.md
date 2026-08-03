# @consultchimps/messages

## 0.2.1

### Patch Changes

- Updated dependencies [6564e24]
  - @consultchimps/core@0.4.0

## 0.2.0

### Minor Changes

- 00740ec: Let every interface reuse the plain-language explanations in its own
  words. `formatHumanResult` and `formatHumanError` now accept an optional
  `{ vocabulary }` option typed as the new exported `MessageVocabulary`, which
  holds the interface-specific phrases: how to retry with overwriting enabled,
  where to find the reference or examples, how to inspect a PowerPoint template
  first, how to point at the created files, and the word for a unit of work.

  Two vocabularies ship with the package. `GENERIC_VOCABULARY` is the new
  default and never names a flag, an executable, or a terminal, so a desktop or
  browser interface can show the guidance unchanged. `CLI_VOCABULARY` reproduces
  the command-line wording, including `--force`, `--help`, and
  `consultchimps pptx inspect-template`.

  Both functions keep their existing signatures, so current callers still
  compile. Library callers that do not pass a vocabulary now receive the neutral
  wording instead of command-line instructions. The `consultchimps` CLI passes
  `CLI_VOCABULARY`, so its output is byte-for-byte unchanged.

## 0.1.1

### Patch Changes

- 5c08c75: Publish typed error-code registries (`FILES_ERRORS`, `PDF_ERRORS`,
  `XLSX_ERRORS`, `PPTX_ERRORS`, with matching `*ErrorCode` unions) so consumers
  can match expected failures without string literals, and make
  `OperationResult` and `OperationPlan` generic over each operation's metric
  names so metric renames become compile-time errors. All runtime values and
  error codes are unchanged; the generics default to `string`, so existing
  consumers keep compiling.
- Updated dependencies [5c08c75]
  - @consultchimps/core@0.3.0

## 0.1.0

### Minor Changes

- c78b35e: Extract the CLI's plain-language result and error rendering into the
  new `@consultchimps/messages` package so desktop and web interfaces can reuse
  the same explanations. The CLI output is unchanged apart from a new recovery
  explanation for cancelled operations.

### Patch Changes

- Updated dependencies [c78b35e]
  - @consultchimps/core@0.2.0
