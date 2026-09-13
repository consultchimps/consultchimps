---
"@consultchimps/core": minor
"@consultchimps/files": minor
"@consultchimps/xlsx": minor
"@consultchimps/db": major
"@consultchimps/messages": minor
"consultchimps": minor
---

Create and reopen persistent SQLite and DuckDB databases, prepare workbook
imports for review, and retain source captures and separately recorded
deliveries. Add bounded workbook reading and file access for large imports.

Replace the database spike's synchronous in-memory API with asynchronous
persistent operations. Applications must migrate to the new database and runtime
entry points. Saved source files are not automatically rewritten.

Add the `consultchimps db` command group. Portable CLI distribution now includes
native engine assets in an archive for its operating system, architecture, and
Node major, rather than a single JavaScript file.

Bind prepared artifact format 3 reviews to their metadata fingerprint and
capture row checksums. Verify consumed row values before committing an import.
Version 1 and 2 spike plans must be regenerated from their source files;
existing working database files keep their format.

Accept leading-plus integer tokens consistently during inference and import.
Keep database and import-plan replacement blocked after failed closure until a
close retry succeeds.

Reject ill-formed Unicode in schema identifiers and Record ID components. Reject
added, missing, or reordered internal metadata columns when opening working
databases and saved import plans.

Preserve supplementary Unicode characters in inferred table names and Record ID
prefixes. Keep streaming workbook cleanup retryable after a scratch-file close
failure, and report failed browser candidate cleanup with recovery locations.

Reject automatic imports with no selected workbook regions before recording an
import. Report retained native staging files and outputs published before a
reopen failure. Keep scratch cleanup retryable and block native replacement
after failed registration cleanup until the unresolved handle closes. Report
successful publication when removing the staging filename fails, and retain
private files when initialization leaves an unclosed engine connection. Reject
SQLite triggers on managed or internal tables before managed writes so they
cannot silently suppress imported rows or alter retry receipts.

Revalidate schema plans before applying changes and reject undeclared reader
cells before capture. Preserve cleanup failures and recovery paths when a native
open or one-step CLI import fails.

Validate internal table types, nullability, and key constraints when opening
databases and saved plans. Delay CLI database results until resource cleanup
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

Reject import receipt replay when the approved plan belongs to another database,
including a copy converted to a different format.

Parse workbook row references in linear time so malformed references cannot
trigger excessive regular-expression backtracking.
