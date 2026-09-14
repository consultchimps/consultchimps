# Database consolidation proposal

Status: slices 1 through 4 implemented and locally verified. The operation is
import, the saved unit of work is a batch, and reusable settings are a profile.
`pnpm check` passed with 1,668 tests and coverage thresholds; the production
browser suite passed 119 tests. Independent standards and specification reviews
reported no actionable findings. Pull request CI and Codex review remain the
external gates. Legacy identity removal is deferred until retained-file
compatibility is known.

## Recommendation

Consolidate the current implementation inside `@consultchimps/db`. Keep explicit
database and prepared-batch handles. Group import operations under
`consultchimps db import`. Give the package sole ownership of profile rules,
saved import reviews, database write outcomes, and internal storage layouts.

The intended result is less state for browser and CLI callers to coordinate. A
browser review should come from the saved batch, and a failed checkpoint should
not make a committed import look safe to repeat. SQLite and DuckDB keep their
own storage and publication rules.

This proposal follows the
[architectural assessment on PR #182](https://github.com/consultchimps/consultchimps/pull/182#issuecomment-5660278960)
and the persistent-file boundary in
[ADR 0005](adr/0005-persistent-database-imports.md). The assessment predates the
latest schema checkpoint fix. That fix is evidence for a shared outcome
contract, not work to repeat.

## What changes for users

The Database tool stays under Online tools at `/tools/db`. Users continue to
create or open a local database, prepare Excel imports, review destinations and
conflicts, apply imports, record batches, and export SQLite or DuckDB files.
Saved reviews reopen without the original Excel files.

The target command inventory is:

```text
consultchimps db create
consultchimps db inspect
consultchimps db schema apply
consultchimps db import prepare
consultchimps db import inspect
consultchimps db import update
consultchimps db import apply
consultchimps db import run
consultchimps db import history
consultchimps db import record
consultchimps db export
```

These are command names, not executable examples with required arguments. The
implementation pass updates help and guides if an option or result changes. It
does not introduce a new package, a `workspace` command, an analytics interface,
or a general operation-pipeline runner.

The visible improvements are consistent naming between callers using the same
policy, saved reviews that agree with the database, and recovery messages that
distinguish completed changes from unfinished saving or cleanup. This
consolidation does not establish a large-workload performance claim by itself.

## Problems grounded in the current implementation

| Current code                                                                                                                                                                     | Duplicated decision                                                                                                         | Proposed owner                                                             |
| -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------- |
| [`workspace.worker.ts`](../apps/docs/src/workers/workspace.worker.ts), `prepareSources`                                                                                          | Calls `draftImportProfile`, discards its return value, then builds another profile and Record ID prefixes                   | `import/profile.ts`                                                        |
| The worker's `HeldImport`, `importDto`, and `heldFromInspection`                                                                                                                 | Keeps a profile, reference, regions, and application state beside persisted plan data; combines these with fresh inspection | `import/inspection.ts`                                                     |
| `ManagedDatabase` and `ManagedImportBatch`                                                                                                                                       | Concurrent, retryable close state                                                                                           | One private close controller                                               |
| Native, browser, CLI, and worker cleanup helpers                                                                                                                                 | Independent close attempts, retained failed owners, and errors after a successful write                                     | A small shared resource collection plus runtime-specific cleanup policy    |
| [`database.ts`](../packages/db/src/database.ts), [`database-layout.ts`](../packages/db/src/internal/database-layout.ts), and [`conversion.ts`](../packages/db/src/conversion.ts) | Internal table columns, constraints, and copy order                                                                         | One private storage descriptor                                             |
| [`prepared.ts`](../packages/db/src/prepared.ts)                                                                                                                                  | Prepared-table DDL and a separate required-layout description                                                               | The same descriptor mechanism                                              |
| [`application-identity.ts`](../packages/db/src/import/application-identity.ts)                                                                                                   | Reconstructs intermediate application identities during ordinary inspection and apply                                       | Strict current identity validation, with a separate compatibility decision |

File size alone is not a reason to split a module. A replacement must remove a
duplicated rule or a sequence callers currently have to remember.

## Package and module boundaries

| Package or module                                        | Responsibility after consolidation                                                                                                |
| -------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------- |
| `@consultchimps/core`                                    | Existing errors, operation results, progress, and cancellation contracts; a narrow resource-close utility if shared by DB and CLI |
| `@consultchimps/files`                                   | Existing input discovery and destination validation                                                                               |
| `@consultchimps/tabular`                                 | Existing runtime-neutral row and column contracts                                                                                 |
| `@consultchimps/xlsx/stream`                             | Existing bounded workbook parsing, reader identity, source leases, and immutable inputs                                           |
| `@consultchimps/db`                                      | Schemas, profiles, captures, applications, batch history, inspections, and write outcomes                                         |
| `@consultchimps/db/node` and `@consultchimps/db/browser` | Native and browser storage, file identity, publication, recovery, and export                                                      |
| `@consultchimps/cli`                                     | Argument parsing, adapter selection, progress, and output                                                                         |
| The docs worker and Database UI                          | Browser files, handles, unsaved form values, protocol messages, and display                                                       |

The proposed file changes are deliberately limited:

- Evolve `packages/db/src/import/profile.ts`, `inspection.ts`, and `types.ts`.
  Do not add parallel `prepareImportReview` or `inspectImportReview` APIs.
- Add `packages/db/src/internal/retryable-close.ts` for the two managed-handle
  classes.
- Add `packages/core/src/resource-close.ts` for cross-package resource ownership
  only. Keep DB persistence semantics out of core.
- Add `packages/db/src/write-completion.ts` for the common database checkpoint
  result. Export this small DB-specific operation and its types from the
  existing package entry point.
- Replace `internal/table-storage.ts` with `internal/storage-schema.ts`, and add
  `internal/storage-layouts.ts` for the concrete metadata definitions.
- Keep native publication and browser backup, restore, and sidecar handling in
  their existing adapter modules unless extraction removes a demonstrated
  dependency.

## Caller-first import design

The following sketches are proposed API changes to existing operations. They are
not currently executable examples. Handle creation and source cleanup are
omitted to focus on the review contract.

```ts
const profile = await draftImportProfile({
  sources,
  naming: { kind: "selection-label" },
});

const captured = await prepareImport({
  database,
  prepared,
  sources,
  profile,
  reviewPage: { limit: 20 },
});
renderReview(captured.inspection);

const resolved = await resolveImport({
  database,
  prepared,
  decisions,
  reviewPage: { limit: 20 },
});
renderReview(resolved.inspection);

const reopened = await inspectImport({
  database,
  prepared: reopenedPreparedHandle,
  page: { limit: 20 },
});
renderReview(reopened);
```

`ImportInspection` remains the authoritative public review type. Extend it with
source display names, destination column information required by the UI, and
persisted application status. `display_name` already exists in prepared captures
and source bindings. Exposing it does not require a new stored column.

Plan-level application state must be explicit before the worker cache is
removed:

```ts
type ImportApplication =
  | { readonly state: "pending" }
  | { readonly state: "applied"; readonly captureIds: readonly string[] };
```

Derive membership from the saved applied batch and validated receipt
relationships, including excluded and reused routes. A route's
`applicationState` alone cannot answer which captures belong to the applied
batch or a later batch.

`prepareImport` and `resolveImport` return the existing operation result or
reference together with an `inspection` when the caller requests a review page.
CLI paths that only need a reference do not pay for a preview. Both operations
use the same internal projection as `inspectImport`.

When inspection is requested, its `inspection.prepared` is the authoritative
returned reference. Build the projection from the final prepared metadata
transaction while holding the operation's existing exclusion. Do not attach a
later `inspectImport` result to an earlier reference. If projection fails after
the plan update, preserve that update outcome and let the caller retry
inspection.

The projection reads one prepared metadata snapshot. Its reference, profile,
bindings, conflicts, and route summaries must agree. Target database state
belongs to a separate file, so the implementation must not claim an atomic
snapshot across both files. Record the observed target revision and detect
changes during inspection. Apply continues to validate the approval fingerprint
and target revision inside its existing write path.

The naming policy has three explicit cases:

```ts
type ImportNaming =
  | { readonly kind: "source-or-selection" }
  | { readonly kind: "selection-label" }
  | { readonly kind: "single-table"; readonly name: string };
```

The CLI keeps its current default and maps `--into` to `single-table`. The
browser requests `selection-label`. Identifier sanitizing, truncation,
reserved-name checks, prefix creation, and collisions have one implementation.
The selection-label policy retains the browser's prefix convention: take the
first eight Unicode code points of the sanitized table name, then uppercase. The
other policies retain word initials. Display labels never replace source
identity keys.

The worker's final `HeldImport` stores the prepared handle and the information
needed to close or remove its browser artifact. Delete its mirrored `profile`,
`ref`, `regions`, and `application` fields. Delete `suggestedTable`, duplicate
prefix generation, and `heldFromInspection`. Convert an `ImportInspection` to a
browser DTO with a pure function. Unsaved user decisions remain UI state until
resolution persists them.

Existing preview pagination remains bounded. Add route keyset pagination in
slice 2, with a default of 50 routes and a maximum of 100 per page. Use a
separate `routePage` option and route cursor so preview cursors keep their
existing meaning. Read bindings with `LIMIT pageSize + 1`, aggregate totals in
SQL, and batch application lookups for the selected route page. Bind cursors to
the plan revision and reject stale cursors. The browser renders and advances
these pages.

This bounds route projection and application lookups, but it does not eliminate
whole-profile JSON validation or the plan-level capture membership list. Measure
those metadata costs separately from captured rows. Do not claim bounded total
metadata memory merely because the DTO's routes and preview rows are paginated.

## Resource ownership and completed writes

Two mechanisms are shared. Their policies stay separate.

`RetryableClose` owns the open, closing, and closed states. Concurrent close
calls share the pending attempt. A failed close retains the engine and
reservation and permits another attempt. A successful close releases ownership
once. Apply this to both managed handle classes.

`OwnedResources` registers closeable resources and attempts independent closes
with `Promise.allSettled`. Successful owners leave the collection; failed owners
remain reachable. Its result identifies failed owners and preserves their
causes. It does not remove files, decide whether publication succeeded, or
translate an error into a warning. Native candidate cleanup, browser replacement
recovery, and CLI input cleanup retain those decisions.

Dependent cleanup remains ordered. An engine must close before its temporary
file can be removed. A retained failed handle must not be hidden inside an
automatic-disposal scope whose caller has no way to retry it.

Database writes need a shared distinction between the SQL transaction and a
requested checkpoint. Add `databaseWrite: "unchanged" | "committed"` to schema,
import, and batch results. Set that field in the transaction branch that knows
what changed. A zero-row import can still commit receipt or batch metadata.

The proposed adapter-facing result and operation are:

```ts
type DatabaseCheckpoint =
  | { readonly state: "checkpoint-completed" }
  | {
      readonly state: "checkpoint-required";
      readonly code: string;
      readonly message: string;
    };

interface DatabaseWriteOutcome<Result> {
  readonly result: Result;
  readonly checkpoint: DatabaseCheckpoint;
}

declare function checkpointDatabaseWrite<
  Result extends DatabaseWriteResult,
>(options: {
  readonly database: Database;
  readonly result: Result;
}): Promise<DatabaseWriteOutcome<Result>>;

interface DatabaseWriteResult {
  readonly databaseWrite: "unchanged" | "committed";
}
```

The result retains its operation-specific receipt and replay information. A
replay does not imply another insertion. Native operations keep their current
transaction policy. The browser passes the successful write result to
`checkpointDatabaseWrite` before emitting its outcome. A completed checkpoint
does not promise that a file survives hardware failure.

The helper contains one checkpoint attempt and returns the operation result even
if that checkpoint needs retry. It never reruns the transaction. An explicit
checkpoint still runs for an unchanged replay because an earlier attempt may
have committed and then failed to checkpoint. Skipping it requires evidence
about persistence that the replay marker alone does not provide.

Pre-commit input and transaction failures still throw the existing structured
errors. The helper catches only the checkpoint call, preserves its cause, and
reports that stage. It does not catch the mutation itself or turn an unexpected
programming error into ordinary success.

Schema application, import application, and batch recording use the same
contract. Display refresh happens separately from the write result. If a summary
query fails after commit, the caller can retry inspection without offering to
repeat the write. The UI clears an applied batch before refreshing display
state.

Artifact publication is a separate event. Native rename, browser replacement,
and export can fail after the public destination exists. Their structured
outcomes retain the publication fact, output identity, and cleanup requirement.
A transaction result cannot describe whether a candidate file was published. No
generic callback-driven lifecycle framework is proposed.

## One definition of internal storage

`DATABASE_STORAGE` and `PREPARED_STORAGE` in `internal/storage-layouts.ts`
describe fixed internal tables. Each table declares ordered column names and
storage types, nullability, primary and unique keys, engine exceptions, and
conversion copy keys and order.

`internal/storage-schema.ts` derives creation SQL, layout validation,
named-column insertion, and conversion metadata from those definitions. It
supports the internal text and integer storage types already in use. User-facing
`TableSchema`, Record ID configuration, relationships, and schema planning
remain separate.

```ts
await createInternalTables(transaction, DATABASE_STORAGE, format);
await insertInternalRow(transaction, DATABASE_STORAGE.tables.database, {
  database_id: id,
  format,
  format_version: BigInt(DATABASE_FILE_FORMAT_VERSION),
  revision: 0n,
});
```

Bulk writes remain batched and use descriptor-derived named columns. There is no
per-row generic object conversion in the capture-row hot path. SQL conflict
clauses remain explicit where that is clearer than adding a query builder.

The deletions are `REQUIRED_DATABASE_SCHEMA`, `REQUIRED_PREPARED_SCHEMA`,
duplicated creation SQL, the internal `COPY_TABLES` column lists, and the
superseded `InternalTableStorage` descriptor. Keep registered user-table
validation and deliberately malformed SQL fixtures independent from the
generator.

Characterization tests must compare actual SQLite and DuckDB layouts with an
independent expected layout. A test that compares generated SQL only with its
own descriptor would miss a shared mistake. Preserve SQLite primary-key
nullability behavior and DuckDB's intentional capture-row primary-key exception.

## Compatibility decisions

Public operation names, TypeScript contracts, CLI commands, and prose use the
accepted vocabulary. Existing `DB_*` error codes remain stable so callers and
support references can continue to identify the same failures. Stored SQL column
names, format discriminators, ID prefixes, and the `.ccplan` extension also
remain unchanged. These compatibility tokens are exceptions to the prose naming
convention.

Preserve working database format 1 and prepared-batch format 3 if
characterization proves the refactor emits their existing layouts. Keep
unsupported prepared formats 1 and 2 untouched and reject them before
current-layout validation. Preserve reader-version provenance and the existing
recapture behavior when the XLSX reader changes.

The raw legacy application-key reconstruction came from intermediate spike work.
The target normal path accepts the current `mapping-v1` identity. Before
removing the fallback, establish whether retained intermediate databases need
it. Published-version evidence alone does not establish which local files exist.

If no retained files need that history, delete the fallback and its legacy-only
tests. If retained files do need it, scope an explicit validated copy migration
before deletion. That operation must leave the source intact, refuse ambiguous
history, and publish only a validated destination. Do not invent provenance or
silently modify files on open.

This compatibility question does not block profile, review, cleanup, or layout
consolidation. No migration command is included in the default scope.

## Implementation sequence and acceptance gates

| Slice                               | Deliverable and deletion                                                                                              | Acceptance gate                                                                                                                                                               |
| ----------------------------------- | --------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 1. Restore correctness baseline     | Validate non-negative baseline revisions at prepared-batch creation and update, before schema or destination mutation | The paused regression passes; invalid creation and overwrite preserve destinations on both engines; valid zero and positive revisions work                                    |
| 2. Authoritative profile and review | Evolve existing DB operations, migrate the worker and CLI, remove mirrored review state and profile builders          | Same-policy naming parity; source-free reopen parity; route edits invalidate approval; same-content inputs retain distinct bindings; lost replies recover without repeat rows |
| 3. Ownership and outcomes           | Share retryable close and retained-owner handling; centralize commit and checkpoint outcomes                          | Failure injection across schema, import, batch, create, and export; failed handles remain retryable; post-commit failures never offer an unsafe repeat                        |
| 4. Internal storage                 | Generate DDL, validation, named writes, and conversion metadata; delete duplicate declarations                        | Independent layout characterization; current files reopen on both engines; corruption detection and conversion round trips retain coverage                                    |
| 5. Compatibility cleanup            | Remove unused intermediate identity reconstruction, or isolate it behind an explicitly scoped migration               | Current identities and receipt replay pass; unsupported files remain untouched; no ambiguous history is accepted                                                              |
| 6. Review and handoff               | Update docs and release intent, run integration gates, push the same PR branch, trigger review                        | Required checks pass on the latest head; review submissions and inline threads are assessed separately; accepted findings have regressions                                    |

Slices 1 and 2 establish the caller contract. Layout implementation can run
alongside slice 3 once the interfaces and file ownership are settled. Agents
must not concurrently edit `prepared.ts`, `database.ts`, or shared import types.
The orchestrator owns those integration points and the outcome semantics;
bounded workers handle callers, layout conversion, and verification.

Slices 1 through 4 now have implementation and focused regression coverage on
this branch. Independent layout tests preserve the existing physical formats.
Slice 5 is deferred until a retained-file inventory or explicit migration is
scoped; the legacy application identity fallback remains readable. Slice 6
remains open until the full package, browser, and cross-platform checks pass
against the final branch head.

## Performance and verification

Preserve bounded workbook reads, batched capture writes, inline fresh checksums,
and the existing bounded scan for reused captures. A receipt-only retry must use
persisted metadata without rereading Excel or scanning observation rows. Review
projection must not materialize captured tables in JavaScript.

A single synthetic 100,000-row run compared the current consolidation branch
with its pre-consolidation baseline. DuckDB prepare time was 340.89 ms before
and 340.79 ms after; apply time was 1,035.71 ms before and 1,029.82 ms after.
SQLite prepare time was 346.37 ms before and 337.08 ms after; apply time was
531.24 ms in both runs. Receipt retries read no source rows in either format.
These single runs exclude XLSX decoding and do not establish browser or large
workload capacity. Cached format 1 databases and format 3 batches reopened on
both engines without changing their file hashes.

Compare the same synthetic workload before and after each affected slice,
recording commit, engine version, row throughput, peak memory, scratch-disk
usage, reopen latency, preview latency, and receipt retry latency. Include a
metadata-heavy case separately from a row-heavy case. Investigate repeatable
regressions; do not trade an additional full capture scan for simpler caller
code.

Export tests verify valid reopened files and conversion results. They also
exercise cancellation, failed close, failed publication, browser backup
restoration, WAL handling, and an occupied destination. Workbook tests retain
the existing missing-value and namespace cases. Consolidation does not justify
another parser rewrite.

Before handoff, run `pnpm check` and the repository's browser integration suite,
then inspect CI for the current PR head. A local test run cannot stand in for
the required cross-platform checks. Read Codex review submissions and inline
threads separately, including new review activity after fixes.

Update the DB package README, library guide, CLI reference, Database tool guide,
relevant Changesets, and PR description where the contract changed. Audit the
existing catalog and navigation for consistency. Update ADR 0005 only after the
proposed decisions are accepted. Documentation must distinguish tested behavior
from performance work that remains unmeasured.

## Alternatives and deferred work

An owned `DatabaseSession` with child `ImportPlanSession` objects would hide
access upgrades and some handle coordination. It would also replace most public
DB operations and introduce ownership and queue rules throughout the adapters.
Persisted approval validation and physical-file exclusion would still be
necessary. That replacement is larger than the demonstrated consolidation need,
so this proposal keeps explicit handles and removes duplicated decisions within
the current operation model.

A cleanup-helper-only refactor is also insufficient. It would leave browser
review authority duplicated and would not solve a committed write being reported
as an ordinary failure. The selected design therefore combines authoritative
operation results with narrow ownership helpers.

Receipt normalization is deferred. Request-to-application membership cannot be
replaced by an application's original request ID because later requests can
reuse that application. Normalization needs a format decision, migration tests,
and a metadata-scale benchmark. It is not an established optimization for
imported rows or analytical joins.

A query planner, new indexes, parallel write framework, analytics UI,
entity-resolution workflow, and generic pipeline executor are outside this
consolidation. Existing captures, applications, batches, schema relationships,
and file-based operation composition remain the foundation for later work.

Reconsider the design if callers still rebuild persisted reviews, a cleanup
abstraction grows engine-specific callback switches, or cross-operation
ownership still requires repeated manual state machines. Reconsider storage
versioning if the descriptor cannot reproduce the current physical layouts.
These are reasons to revise the design before adding another layer.

## Principles that changed the proposal

- **Foundational thinking** led to defining the inspection and write outcome
  contracts before moving worker code.
- **Subtract before you add** led to deleting mirrored worker state and
  duplicate storage declarations in the same slices that replace them.
- **Minimize reader load** led to one operation API instead of a second review
  API or session facade.
- **Model the domain** kept internal metadata storage separate from user schemas
  and kept captures, applications, and batches distinct.
- **Make operations idempotent** kept receipt replay authoritative and excluded
  automatic transaction retries after checkpoint failure.
- **Prove it works** led to independent layout fixtures, real file reopen tests,
  browser failure injection, and measured performance comparisons.
- **Sequence work into verifiable units** separated the correctness baseline,
  caller migration, cleanup, layouts, and conditional compatibility removal.
