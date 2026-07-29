# @consultchimps/pptx

## 0.2.0

### Minor Changes

- 6870be2: Add local PowerPoint template inspection and text population from
  selected Excel worksheet records, with formatting-preserving slide cloning,
  complete pre-write validation, safe overwrite handling, and detailed CLI
  guidance. Also make absolute file-glob discovery portable on Windows.
- 2c4065f: Support PowerPoint placeholders split across adjacent text runs and
  default PowerPoint population to the first template slide and first worksheet
  when those selections are omitted.

### Patch Changes

- Updated dependencies [6870be2]
- Updated dependencies [2c4065f]
  - @consultchimps/xlsx@0.4.0
  - @consultchimps/files@0.1.1
