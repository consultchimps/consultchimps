# @consultchimps/db

## 1.0.0

### Major Changes

- 6606d1e: Group staged imports under `db import prepare`, `inspect`, `update`,
  and `apply`. Use `db import run` for one-step execution, `--profile` for
  reusable settings, and `--batch` for a saved review. Rename the library
  contracts to import profiles and batches while preserving the current stored
  database and batch formats.

  Page saved import routes independently from row previews and expose applied
  capture membership in import inspection. Share destination naming policies
  between callers while preserving their Record ID prefix conventions. Count
  shared captures once in review totals while retaining each source binding.

  Report whether database writes committed or returned an unchanged result. Keep
  completed write results available after checkpoint or summary-refresh
  failures. Add retryable resource ownership through `OwnedResources` in the
  core package.

  Create and reopen persistent SQLite and DuckDB databases, prepare workbook
  imports for review, and retain source captures and separately recorded
  batches. Add bounded workbook reading and file access for large imports.

  Replace the database spike's synchronous in-memory API with asynchronous
  persistent operations. Applications must migrate to the new database and
  runtime entry points. Saved source files are not automatically rewritten.

  Add the `consultchimps db` command group. Portable CLI distribution now
  includes native engine assets in an archive for its operating system,
  architecture, and Node major, rather than a single JavaScript file.

  Bind prepared artifact format 3 reviews to their metadata fingerprint and
  capture row checksums. Verify consumed row values before committing an import.
  Version 1 and 2 spike batches must be regenerated from their source files;
  existing working database files keep their format.

  Accept leading-plus integer tokens consistently during inference and import.
  Keep database and import-batch replacement blocked after failed closure until
  a close retry succeeds.

  Reject ill-formed Unicode in schema identifiers and Record ID components.
  Reject added, missing, or reordered internal metadata columns when opening
  working databases and saved import batches.

  Preserve supplementary Unicode characters in inferred table names and Record
  ID prefixes. Keep streaming workbook cleanup retryable after a scratch-file
  close failure, and report failed browser candidate cleanup with recovery
  locations.

  Reject automatic imports with no selected workbook regions before recording an
  import. Report retained native staging files and outputs published before a
  reopen failure. Keep scratch cleanup retryable and block native replacement
  after failed registration cleanup until the unresolved handle closes. Report
  successful publication when removing the staging filename fails, and retain
  private files when initialization leaves an unclosed engine connection. Reject
  SQLite triggers on managed or internal tables before managed writes so they
  cannot silently suppress imported rows or alter retry receipts.

  Revalidate schema plans before applying changes and reject undeclared reader
  cells before capture. Preserve cleanup failures and recovery paths when a
  native open or one-step CLI import fails.

  Validate internal table types, nullability, and key constraints when opening
  databases and saved batches. Delay CLI database results until resource cleanup
  finishes, preserving valid JSON and recovery details after a committed
  operation. Check cancellation during browser replacement copies and preserve
  failed browser open cleanup so live owners continue to block replacement.

  Reject repeated workbook-property declarations and duplicate import source or
  selection keys before capture. Reject stored application counts that disagree
  with their capture before reporting reuse or recording new retry receipts.

  Reject browser automatic imports when no workbook regions are selected. Check
  metadata allocation counters when opening a database and row allocation before
  new table applications. Report structured corruption errors for counter
  conflicts and pending Record ID collisions while preserving valid counter gaps
  and the already-applied retry path.

  Reject import receipt replay when the approved batch belongs to another
  database, including a copy converted to a different format.

  Parse workbook row references in linear time so malformed references cannot
  trigger excessive regular-expression backtracking.

  Detect DuckDB view changes before applying reviewed schema or import batches.
  Validate saved import receipt totals against application metadata before
  returning a successful retry result.

  Read the workbook date system only from its SpreadsheetML property element so
  extension metadata cannot shift imported dates.

  Reject saved import batches containing rows from an unfinished capture when
  reopening them, without deleting the batch or scanning row values.

  Ignore foreign or misplaced workbook sheet and named-range declarations.
  Reject retry receipts that substitute applications from unrelated captures or
  tables while retaining reuse of applications from earlier requests.

  Apply namespace and ancestry checks to worksheet values, shared strings,
  styles, and table references. Advance the workbook reader version so new
  preparations read files again when only an older reader capture exists.
  Existing imported rows and capture history are not rewritten automatically.

  Reject unsupported schema and import-profile versions from JavaScript callers
  before publishing database or prepared-batch outputs.

  Reject unsupported XML document roots instead of reporting an empty workbook.
  Verify batch capture memberships before returning an import retry receipt.

  Reject worksheet aliases and require supported OOXML relationship role URIs.
  Preserve orphaned browser DuckDB recovery files by refusing publication over
  incomplete storage.

  Report committed batch updates when preparation checkpointing fails.
  Distinguish caller-owned batches from private staging discarded before
  publication. Preserve DuckDB appender operation and cleanup failures together,
  and retry only the remaining native resources after a partially failed
  database close.

  Preserve the original workbook initialization error when lazy opening fails.
  Closing a source with no acquired stream no longer creates a false cleanup
  failure or blocks later preparations.

  Preserve transaction and rollback failures together in SQLite and DuckDB,
  across native and browser runtimes. Block further operations on an unresolved
  connection while leaving cleanup available, and retain the original error when
  rollback succeeds.

### Minor Changes

- 467f75c: Add `@consultchimps/db/sqlite-read` for temporary read-only SQLite
  catalogs in Node and browser workers. The entry uses the official SQLite WASM
  runtime, preserves 64-bit integers and binary values, provides configurable
  read limits, and creates no persistent workspace files.
- 4bc99d2: Raise the supported Node.js floor from 22.0.0 to 22.14.0. Continuous
  integration now tests that exact version: pnpm 11 requires 22.13, and
  better-sqlite3 13 crashes on the 22.13 patch line, so 22.14.0 is the lowest
  release on which every published package works.

### Patch Changes

- Updated dependencies [4bc99d2]
- Updated dependencies [6606d1e]
- Updated dependencies [04dd6de]
  - @consultchimps/core@0.6.0
  - @consultchimps/files@0.5.0
  - @consultchimps/xlsx@0.18.0
