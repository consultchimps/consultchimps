# @consultchimps/xlsx

## Streaming workbook reads

The `@consultchimps/xlsx/stream` entry reads worksheet, named-table, and range
selections from random-access OOXML bytes. A shared workbook session yields
bounded row batches and uses scratch files for shared strings. It retains
numeric tokens and distinguishes formula caches, missing caches, and errors. It
reads date styles and both workbook date systems without loading the entire
worksheet into JavaScript arrays.

The reader uses namespace-aware paths for worksheet values, shared strings,
styles, workbook declarations, and table references. Foreign extension elements
do not become spreadsheet data merely because they use the same element names.
Unsupported document roots and namespaces return `XLSX_READ_FAILED`.
Relationship roles must use supported Transitional or Strict OOXML URIs.
Workbook sheets that reference the same worksheet part are rejected before
selection, so aliases cannot import the same physical rows twice.

Repeated workbook-property declarations are rejected with `XLSX_READ_FAILED`,
including declarations with matching values, so a later declaration cannot
silently replace the date system.

Close region readers and their workbook session after use. If closure fails,
resolve the storage issue and call `close()` again on the same reader or
session. Concurrent close calls share a pending attempt, and a retry closes only
the resources whose cleanup has not succeeded. `XLSX_CLEANUP_FAILED` identifies
a failed cleanup attempt. If opening or reading also fails, the error retains
the operation and cleanup causes; resolve the reported storage issue before
retrying the operation.

See the
[library guide](https://consultchimps.github.io/consultchimps/docs/libraries/).
