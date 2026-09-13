# @consultchimps/xlsx

## Streaming workbook reads

The `@consultchimps/xlsx/stream` entry reads worksheet, named-table, and range
selections from random-access OOXML bytes. A shared workbook session yields
bounded row batches and uses scratch files for shared strings. It retains
numeric tokens and distinguishes formula caches, missing caches, and errors. It
reads date styles and both workbook date systems without loading the entire
worksheet into JavaScript arrays.

Close region readers and their workbook session after use. If closure fails,
resolve the storage issue and call `close()` again on the same reader or
session. Concurrent close calls share a pending attempt, and a retry closes only
the resources whose cleanup has not succeeded. `XLSX_CLEANUP_FAILED` identifies
a failed cleanup attempt. If opening or reading also fails, the error retains
the operation and cleanup causes; resolve the reported storage issue before
retrying the operation.

See the
[library guide](https://consultchimps.github.io/consultchimps/docs/libraries/).
