# Storage experiment results

Measured on 2026-09-12 on an Apple M4 Pro with 24 GiB RAM and 12 logical CPUs.
Node was 24.18.0 and Chromium was 151.0.7922.34. DuckDB-Wasm ran engine 1.5.4
with one worker. Native DuckDB used 1.5.5 and four threads for the format
comparison. Both used a 1 GB database memory setting.

These are development measurements, not a supported-capacity promise. Other
checks ran on the same machine during parts of the experiment. The fixtures are
narrow and repetitive. First and second executions are not cold-cache and
warm-cache measurements.

## Browser storage and artifact

The [storage result](results/storage.json) creates rows inside DuckDB. It does
not parse Excel.

| Check                                         | Measured result                             |
| --------------------------------------------- | ------------------------------------------- |
| Insert the first 300,000 rows                 | 117 ms                                      |
| Append to seven million rows                  | 1,899 ms                                    |
| Append to seventy million rows                | 17,123 ms                                   |
| Filtered dataset join and decimal calculation | 59 ms, then 37 ms                           |
| Reloaded database count                       | 70,000,000 rows, 7,000,000 CDE flags        |
| Exported database                             | 432,812,032 bytes                           |
| Stream the export to the local test server    | 329 ms                                      |
| Native reopen                                 | Counts and subsequent curation row verified |

The join filters one source file, selects seven million observations, joins
100,000 dataset records, and returns twenty domain groups. This measures one
specific many-to-one join, not the performance of arbitrary vendor queries.

The first OPFS attempt created tables that disappeared on reload. Explicit
registration of the database and WAL file handles with direct I/O passed the
reopen checks. That result establishes the tested graceful-close behavior. Crash
recovery, transaction interruption, quota failures, browser eviction, and
conflicting writers remain unverified.

## Excel parsing

The [small Excel result](results/excel-small.json) imports 300,000 rows from a
real generated workbook in 3,250 ms. It verifies 30,000 CDE flags after page
reload and native reopen. The resulting database is 2,109,440 bytes.

The existing ConsultChimps worksheet reader took 8,159 ms and reached about 1.96
GB maximum resident memory for the same 300,000-row source. This separate Node
measurement used the existing built xlsx byte API with a 2 GB JavaScript heap
limit. It is a baseline for the current whole-workbook reader, not a
browser-versus-native engine comparison.

The [seventy-million-row parsing run](results/excel-seventy-million.json) spent
873,917 ms, about 14.6 minutes, in its ten workbook parsing and insertion loops
and exported a 422,850,560-byte database. This excludes hashing, plan review,
final export, and application-level provenance work. Native DuckDB independently
[verified](results/excel-seventy-million-native.json) 70,000,000 observations,
7,000,000 reported CDE flags, and ten source scopes.

The run parsed one seven-million-row workbook ten times under different source
scope labels. The workbook is 170,663,567 bytes and contains seven sheets of
repeated narrow data. This deliberately bypasses deduplication to stress Excel
parsing. It is not a test of ten distinct vendor files or import receipts.

The [memory sample](results/excel-memory.json) reached about 4.0 GiB for the
Node server and Chromium process tree. Sampling began after the first two sheets
and ran every two seconds, so this is a sampled maximum over part of the run. It
is not a browser memory guarantee. The 1 GB database setting did not cap the
whole process group at 1 GB.

This long run used the exploratory runner before the checked-in harness was
consolidated. The checked-in `excel` command repeats its source-reading sequence
and adds assertions. The checked-in small-import and storage modes were run with
their assertions enabled. Repeating the consolidated long run is still required
before using it as an automated release check.

## DuckDB storage versus SQLite compatibility

The [format comparison](results/comparison.json) uses seven million typed
observations and 100,000 dataset records. DuckDB executes the query in both
cases. The filtered join returns 700,000 observations grouped into twenty rows.

| Working format                         | First measured join | Second measured join |
| -------------------------------------- | ------------------- | -------------------- |
| Native DuckDB                          | 6 ms                | 3 ms                 |
| SQLite through DuckDB, without indexes | 153 ms              | 152 ms               |
| SQLite through DuckDB, with indexes    | 153 ms              | 156 ms               |

The SQLite indexes cover source-file filtering and the dataset key. They did not
improve this DuckDB query path. Native DuckDB occupies 39,858,176 bytes. SQLite
including these indexes occupies 367,173,632 bytes.

DuckDB's SQLite exporter stored the decimal amount column as text. The measured
query explicitly converts amounts to `DECIMAL(18,2)` in both cases and verifies
a known result. A compatibility export therefore needs declared type mappings.
It cannot promise unchanged schemas and calculations just because both formats
accept SQL.

## Engineering decision

Recommend native DuckDB for the new analytical working file, subject to the
pending product choice. Keep the existing SQLite workspace behavior while the
new path is developed. Make SQLite conversion an explicit snapshot export.

The next implementation work should establish bounded source reading, persistent
import receipts, immutable source observations, and file-attributed queries. The
current in-memory SQLite and whole-workbook reader are useful references for
semantics, but they are not the large-import runtime.

The DuckDB Excel extension supplies a useful performance baseline. It still
needs conformance checks for physical source rows, table regions, formula
caches, errors, dates, leading-zero identifiers, and shared strings before
adoption. The browser package's prerelease status also needs a production
release decision.

The experiment does not implement duplicate prevention, saved mappings, DQ,
current-inventory selection, relationships, CDE approval, or a user-facing
import operation. Those remain in the staged plan. No public operation is marked
available by this change.
