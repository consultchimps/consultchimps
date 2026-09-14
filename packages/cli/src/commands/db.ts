import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import {
  ConsultChimpsError,
  type OperationControlOptions,
} from "@consultchimps/core";
import {
  applyImport,
  applySchema,
  draftImportProfile,
  inspectDatabase,
  inspectImport,
  listBatches,
  parseDatabaseSchema,
  parseBatchContext,
  parseImportProfile,
  planSchema,
  planConversion,
  prepareImport,
  recordBatch,
  replaceImportProfile,
  resolveImport,
  type DatabaseFormat,
  type ImportBatchRef,
  type ReadyImportBatchRef,
} from "@consultchimps/db";
import {
  createDatabase,
  createImportBatch,
  exportDatabase,
  inspectFileKind,
  openDatabase,
  openImportBatch,
  prepareImportFile,
} from "@consultchimps/db/node";
import { planFilePublication } from "@consultchimps/files";
import { Option, type Command } from "commander";

import {
  openDbInputs,
  readDbDocument,
  type DbInputOptions,
} from "../db-inputs.js";
import {
  formatConversionPlan,
  formatDatabaseInspection,
  formatDeliveryPage,
  formatImportInspection,
  formatImportResolution,
  formatSchemaPlan,
} from "../db-report.js";
import {
  finishCliImport,
  type CliImportOutcome,
} from "../db-import-cleanup.js";
import {
  type DbCommandOutput,
  withDeferredDbCommandOutput,
} from "../db-command-output.js";
import { createCliProgress } from "../progress.js";
import { withoutTerminalControlsInProse } from "../text.js";

interface ImportOptions extends DbInputOptions {
  output?: string;
  profile?: string;
  into?: string;
  force?: boolean;
  context?: string;
  requestId?: string;
}

function requireReady(
  prepared: ImportBatchRef | ReadyImportBatchRef,
): ReadyImportBatchRef {
  if (prepared.state !== "ready") {
    throw new ConsultChimpsError(
      "DB_IMPORT_NEEDS_REVIEW",
      "The batch has unresolved table or column conflicts. Use db import prepare to save the captured data, db import inspect to review it, and db import update with a corrected profile before applying.",
    );
  }
  return prepared;
}

function collect(value: string, previous: string[] = []): string[] {
  return [...previous, value];
}

function chooseFormat(
  output: string,
  format?: string,
  fallback?: DatabaseFormat,
): DatabaseFormat {
  const extension = path.extname(output).toLowerCase();
  const inferred =
    extension === ".duckdb"
      ? "duckdb"
      : extension === ".sqlite" || extension === ".sqlite3"
        ? "sqlite"
        : undefined;
  if (format !== undefined && format !== "sqlite" && format !== "duckdb") {
    throw new ConsultChimpsError(
      "DB_INVALID_FORMAT",
      "Choose sqlite or duckdb as the database format.",
    );
  }
  if (format && inferred && inferred !== format) {
    throw new ConsultChimpsError(
      "DB_FORMAT_CONFLICT",
      "The filename extension and chosen database format disagree. Change one before retrying.",
    );
  }
  const selected = format ?? inferred ?? fallback;
  if (!selected)
    throw new ConsultChimpsError(
      "DB_FORMAT_REQUIRED",
      "Specify --format sqlite or --format duckdb when the filename does not identify its format.",
    );
  return selected;
}

async function withControls<T>(
  output: DbCommandOutput,
  work: (controls: OperationControlOptions) => Promise<T>,
): Promise<T> {
  const progress = createCliProgress(output.json());
  const controller = new AbortController();
  const cancel = (): void => controller.abort();
  process.once("SIGINT", cancel);
  try {
    return await work({
      signal: controller.signal,
      onProgress: progress.report,
    });
  } finally {
    process.removeListener("SIGINT", cancel);
    progress.finish();
  }
}

function sources(command: Command): Command {
  return command
    .requiredOption(
      "--input <source>",
      "workbook path or alias=path; existing paths are used as written",
      collect,
    )
    .option(
      "--profile <file>",
      "reusable JSON table routing and column mapping",
    )
    .option("--sheet <name>", "select one worksheet from one workbook")
    .option("--table <name>", "select one named Excel Table")
    .option(
      "--range <reference>",
      "select one named range or worksheet A1 rectangle",
    )
    .option(
      "--header-row <number>",
      "worksheet header row, starting at 1",
      Number,
      1,
    )
    .option("--hidden", "include hidden worksheets in default sheet selection")
    .option("--into <table>", "destination name for one selected region");
}

async function prepare(
  databasePath: string,
  options: ImportOptions,
  output: DbCommandOutput,
  apply: boolean,
): Promise<void> {
  await withDeferredDbCommandOutput(
    output,
    async (output) => {
      await withControls(output, async (controls) => {
        const delivery = options.context
          ? parseBatchContext(await readDbDocument(options.context))
          : undefined;
        const database = await openDatabase({ path: databasePath });
        let temporary: string | undefined;
        let inputs: Awaited<ReturnType<typeof openDbInputs>> | undefined;
        let prepared: Awaited<ReturnType<typeof createImportBatch>> | undefined;
        let outcome: CliImportOutcome = { status: "completed" };
        try {
          const suppliedProfile = options.profile
            ? parseImportProfile(await readDbDocument(options.profile))
            : undefined;
          inputs = await openDbInputs(options, controls, suppliedProfile);
          const sourceList = inputs.workbooks.map(
            (workbook) => workbook.source,
          );
          const profile =
            suppliedProfile ??
            (await draftImportProfile({
              sources: sourceList,
              naming:
                options.into === undefined
                  ? { kind: "source-or-selection" }
                  : { kind: "single-table", name: options.into },
            }));
          const inspected = await inspectDatabase({ database });
          if (!options.output)
            temporary = await mkdtemp(path.join(tmpdir(), "cc-import-plan-"));
          const planPath =
            options.output ?? path.join(temporary ?? "", "review.ccplan");
          const protectedInputPaths = [
            ...inputs.paths,
            ...(options.profile ? [options.profile] : []),
            ...(options.context ? [options.context] : []),
          ];
          if (!apply) {
            const preparedFile = await prepareImportFile({
              path: planPath,
              database,
              sources: sourceList,
              profile,
              baselineRevision: inspected.revision,
              overwrite: options.force,
              protectedInputPaths,
              ...controls,
            });
            output.result(preparedFile.result);
            output.prose(
              "Review this batch with db import inspect, then apply it with db import apply. The saved batch can contain source values; keep it private.\n",
            );
          } else {
            prepared = await createImportBatch({
              path: planPath,
              database,
              profile,
              baselineRevision: inspected.revision,
              overwrite: options.force,
              protectedInputPaths,
            });
            const preparedOutcome = await prepareImport({
              database,
              prepared,
              sources: sourceList,
              profile,
              ...controls,
            });
            const approved = requireReady(
              preparedOutcome.prepared.state === "ready"
                ? preparedOutcome.prepared
                : await resolveImport({ database, prepared, decisions: [] }),
            );
            output.result(
              await applyImport({
                database,
                prepared,
                approved,
                requestId: options.requestId ?? prepared.id,
                batchContext: delivery,
                ...controls,
              }),
            );
          }
        } catch (error) {
          outcome = { status: "failed", error };
        }
        const ownedPrepared = prepared;
        const ownedInputs = inputs;
        await finishCliImport({
          outcome,
          temporaryPath: temporary,
          ...(ownedPrepared === undefined
            ? {}
            : { closePrepared: () => ownedPrepared.close() }),
          ...(ownedInputs === undefined
            ? {}
            : { closeInputs: () => ownedInputs.close() }),
          closeDatabase: () => database.close(),
          removeTemporary: () =>
            temporary === undefined
              ? Promise.resolve()
              : rm(temporary, { recursive: true, force: true }),
        });
      });
    },
    apply
      ? [databasePath]
      : options.output === undefined
        ? []
        : [options.output],
  );
}

export function registerDbCommands(
  program: Command,
  output: DbCommandOutput,
): void {
  const db = program
    .command("db")
    .description(
      "Create persistent local databases, manage schemas, and import workbook submissions",
    )
    .addHelpText(
      "after",
      "\nStart with: consultchimps db create -o inventory.duckdb\nThen: consultchimps db import prepare inventory.duckdb --input inventory.xlsx -o review.ccplan\nReview: consultchimps db import inspect review.ccplan\nApply: consultchimps db import apply inventory.duckdb --batch review.ccplan\n",
    );

  db.command("create")
    .description("Create a persistent SQLite or DuckDB file")
    .requiredOption("-o, --output <file>", "new database file")
    .addOption(
      new Option("--format <format>", "database storage format").choices([
        "sqlite",
        "duckdb",
      ]),
    )
    .option("--schema <file>", "versioned JSON table definitions")
    .option("-f, --force", "allow replacement of an existing output")
    .addHelpText(
      "after",
      "\nExample: consultchimps db create -o inventory.sqlite --schema schema.json\n",
    )
    .action(
      async (options: {
        output: string;
        format?: string;
        schema?: string;
        force?: boolean;
      }) => {
        await withDeferredDbCommandOutput(
          output,
          async (output) => {
            const schema = options.schema
              ? parseDatabaseSchema(await readDbDocument(options.schema))
              : undefined;
            await planFilePublication({
              output: options.output,
              inputs: options.schema ? [options.schema] : [],
              overwrite: options.force,
            });
            const created = await createDatabase({
              path: options.output,
              format: chooseFormat(options.output, options.format),
              schema,
              overwrite: options.force,
            });
            try {
              output.result(created.result);
            } finally {
              await created.database.close();
            }
          },
          [options.output],
        );
      },
    );

  db.command("inspect")
    .description("Inspect a database without changing it")
    .argument("<database>", "SQLite or DuckDB database file")
    .action(async (file: string) => {
      await withDeferredDbCommandOutput(output, async (output) => {
        const kind = await inspectFileKind({ path: file });
        if (kind.kind === "unmanaged-database") {
          const tables = kind.tables
            .map(
              (table) =>
                `${withoutTerminalControlsInProse(table.name)}: ${table.columns.length} columns\n`,
            )
            .join("");
          output.data(
            kind,
            `${kind.format === "duckdb" ? "DuckDB" : "SQLite"} database, read-only inspection\nThis file has no ConsultChimps import history. No tables or metadata were added.\n${tables}Create a separate ConsultChimps database for managed imports.\n`,
          );
          return;
        }
        if (kind.kind === "prepared-import") {
          throw new ConsultChimpsError(
            "DB_EXPECTED_DATABASE",
            "This file is a saved import batch. Inspect it with db import inspect.",
          );
        }
        const database = await openDatabase({ path: file, readonly: true });
        try {
          const inspection = await inspectDatabase({ database });
          output.data(inspection, formatDatabaseInspection(inspection));
        } finally {
          await database.close();
        }
      });
    });

  db.command("schema")
    .description("Manage database table definitions")
    .command("apply")
    .description("Review or apply additive schema changes")
    .argument("<database>")
    .requiredOption("--file <schema>", "versioned JSON schema")
    .option("--dry-run", "report proposed changes without applying them")
    .action(
      async (
        databasePath: string,
        options: { file: string; dryRun?: boolean },
      ) => {
        await withDeferredDbCommandOutput(
          output,
          async (output) => {
            const schema = parseDatabaseSchema(
              await readDbDocument(options.file),
            );
            const database = await openDatabase({
              path: databasePath,
              readonly: options.dryRun === true,
            });
            try {
              const plan = await planSchema({ database, schema });
              if (options.dryRun) output.data(plan, formatSchemaPlan(plan));
              else output.result(await applySchema({ database, plan }));
            } finally {
              await database.close();
            }
          },
          [databasePath],
        );
      },
    );

  const importCommand = db
    .command("import")
    .description("Prepare, review, apply, and audit workbook batches");

  importCommand
    .command("inspect")
    .description("Inspect a saved batch without reading Excel")
    .argument("<batch>", "saved batch file")
    .option("--limit <number>", "maximum preview rows", Number, 20)
    .option(
      "--cursor <cursor>",
      "preview cursor returned by a prior batch inspection",
    )
    .option("--route-limit <number>", "maximum routes", Number, 50)
    .option(
      "--route-cursor <cursor>",
      "route cursor returned by a prior batch inspection",
    )
    .option(
      "--database <file>",
      "target database containing reused rows for a saved batch preview",
    )
    .action(
      async (
        batch: string,
        options: {
          limit: number;
          cursor?: string;
          database?: string;
          routeLimit: number;
          routeCursor?: string;
        },
      ) => {
        await withDeferredDbCommandOutput(output, async (output) => {
          const prepared = await openImportBatch({
            path: batch,
            readonly: true,
          });
          try {
            const database =
              options.database === undefined
                ? undefined
                : await openDatabase({
                    path: options.database,
                    readonly: true,
                  });
            try {
              const inspection = await inspectImport({
                database,
                prepared,
                page: { limit: options.limit, cursor: options.cursor },
                routePage: {
                  limit: options.routeLimit,
                  cursor: options.routeCursor,
                },
              });
              output.data(inspection, formatImportInspection(inspection));
            } finally {
              await database?.close();
            }
          } finally {
            await prepared.close();
          }
        });
      },
    );

  sources(
    importCommand
      .command("prepare")
      .description("Capture workbook data into a durable, reviewable batch")
      .argument("<database>"),
  )
    .requiredOption("-o, --output <file>", "private saved batch file")
    .option("-f, --force", "allow replacing an existing batch output")
    .action((database: string, options: ImportOptions) =>
      prepare(database, options, output, false),
    );

  sources(
    importCommand
      .command("run")
      .description("Prepare and apply a workbook batch with one profile")
      .argument("<database>"),
  )
    .option(
      "--context <file>",
      "batch label, scope, and reported attributes as JSON",
    )
    .option("--request-id <id>", "retry key for this application")
    .action((database: string, options: ImportOptions) =>
      prepare(database, options, output, true),
    );

  importCommand
    .command("apply")
    .description("Apply a reviewed saved batch")
    .argument("<database>")
    .requiredOption("--batch <file>", "saved batch file")
    .option("--context <file>", "batch context JSON")
    .option("--request-id <id>", "retry key for this application")
    .action(
      async (
        databasePath: string,
        options: { batch: string; context?: string; requestId?: string },
      ) => {
        await withDeferredDbCommandOutput(
          output,
          async (output) => {
            await withControls(output, async (controls) => {
              const delivery = options.context
                ? parseBatchContext(await readDbDocument(options.context))
                : undefined;
              const database = await openDatabase({ path: databasePath });
              try {
                let prepared = await openImportBatch({
                  path: options.batch,
                  readonly: true,
                });
                try {
                  let review = await inspectImport({
                    database,
                    prepared,
                    page: { limit: 1 },
                  });
                  if (review.prepared.state !== "ready") {
                    await prepared.close();
                    prepared = await openImportBatch({ path: options.batch });
                    review = await inspectImport({
                      database,
                      prepared,
                      page: { limit: 1 },
                    });
                  }
                  const approved = requireReady(
                    review.prepared.state === "ready"
                      ? review.prepared
                      : await resolveImport({
                          database,
                          prepared,
                          decisions: [],
                        }),
                  );
                  output.result(
                    await applyImport({
                      database,
                      prepared,
                      approved,
                      batchContext: delivery,
                      requestId: options.requestId ?? prepared.id,
                      ...controls,
                    }),
                  );
                } finally {
                  await prepared.close();
                }
              } finally {
                await database.close();
              }
            });
          },
          [databasePath],
        );
      },
    );

  importCommand
    .command("update")
    .description("Update a saved batch's table routing and column mapping")
    .argument("<database>")
    .requiredOption("--batch <file>", "saved batch file")
    .option(
      "--profile <file>",
      "replacement import profile; omitted routes stop table loading; omit this option to re-review captured data",
    )
    .action(
      async (
        databasePath: string,
        options: { batch: string; profile?: string },
      ) => {
        await withDeferredDbCommandOutput(
          output,
          async (output) => {
            const profile = options.profile
              ? parseImportProfile(await readDbDocument(options.profile))
              : undefined;
            const database = await openDatabase({ path: databasePath });
            try {
              const prepared = await openImportBatch({ path: options.batch });
              try {
                const resolved =
                  profile === undefined
                    ? await resolveImport({
                        database,
                        prepared,
                        decisions: [],
                        rebase: true,
                      })
                    : await replaceImportProfile({
                        database,
                        prepared,
                        profile,
                        rebase: true,
                      });
                output.completedData(
                  resolved,
                  formatImportResolution(resolved),
                );
              } finally {
                await prepared.close();
              }
            } finally {
              await database.close();
            }
          },
          [options.batch],
        );
      },
    );

  importCommand
    .command("history")
    .description("List recorded batches and their source captures")
    .argument("<database>")
    .option("--limit <number>", "maximum batch records", Number, 50)
    .option("--cursor <cursor>", "pagination cursor from a prior response")
    .action(
      async (
        databasePath: string,
        options: { limit: number; cursor?: string },
      ) => {
        await withDeferredDbCommandOutput(output, async (output) => {
          const database = await openDatabase({
            path: databasePath,
            readonly: true,
          });
          try {
            const deliveries = await listBatches({ database, ...options });
            output.data(deliveries, formatDeliveryPage(deliveries));
          } finally {
            await database.close();
          }
        });
      },
    );

  importCommand
    .command("record")
    .description("Record another batch without importing row values again")
    .argument("<database>")
    .requiredOption(
      "--capture <id>",
      "capture ID, repeat for additional captures",
      collect,
    )
    .requiredOption("--context <file>", "batch context JSON")
    .requiredOption(
      "--request-id <id>",
      "stable key used to retry this batch record safely",
    )
    .action(
      async (
        databasePath: string,
        options: { capture: string[]; context: string; requestId: string },
      ) => {
        await withDeferredDbCommandOutput(
          output,
          async (output) => {
            const context = parseBatchContext(
              await readDbDocument(options.context),
            );
            const database = await openDatabase({ path: databasePath });
            try {
              output.result(
                await recordBatch({
                  database,
                  captureIds: options.capture,
                  context,
                  requestId: options.requestId,
                }),
              );
            } finally {
              await database.close();
            }
          },
          [databasePath],
        );
      },
    );

  db.command("export")
    .description(
      "Create a validated database copy or convert its storage format",
    )
    .argument("<database>")
    .requiredOption("-o, --output <file>", "independent database output file")
    .addOption(
      new Option("--format <format>", "output storage format").choices([
        "sqlite",
        "duckdb",
      ]),
    )
    .option(
      "--dry-run",
      "review format changes and unsupported objects without writing",
    )
    .option("-f, --force", "allow replacing an existing output")
    .addHelpText(
      "after",
      "\nExample: consultchimps db export inventory.sqlite -o inventory.duckdb\nReview first: consultchimps db export inventory.sqlite -o inventory.duckdb --dry-run\n",
    )
    .action(
      async (
        databasePath: string,
        options: {
          output: string;
          format?: string;
          dryRun?: boolean;
          force?: boolean;
        },
      ) => {
        await withDeferredDbCommandOutput(
          output,
          async (output) => {
            if (options.dryRun !== true) {
              await planFilePublication({
                output: options.output,
                inputs: [databasePath],
                overwrite: options.force,
              });
            }
            await withControls(output, async (controls) => {
              const database = await openDatabase({
                path: databasePath,
                readonly: options.dryRun === true,
              });
              try {
                const format = chooseFormat(
                  options.output,
                  options.format,
                  database.format,
                );
                if (options.dryRun) {
                  const plan = await planConversion({ database, format });
                  output.data(plan, formatConversionPlan(plan));
                } else
                  output.result(
                    await exportDatabase({
                      database,
                      output: options.output,
                      format,
                      overwrite: options.force,
                      ...controls,
                    }),
                  );
              } finally {
                await database.close();
              }
            });
          },
          [options.output],
        );
      },
    );
}
