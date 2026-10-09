/**
 * The worker that performs every byte-level operation the tool pages offer.
 *
 * Splitting a large PDF or rewriting a workbook package is seconds of tight,
 * synchronous work. Running it here keeps the tab responsive: the page only
 * hands over the chosen `File`s, renders progress events, and receives the
 * outputs as `Blob`s. Reading each input happens here, never in the page.
 *
 * One worker serves every format, and each engine is still pulled in with a
 * dynamic `import()` on the first task that needs it, so opening a tool page
 * downloads no engine until someone asks for a preview or a run.
 */
import {
  ConsultChimpsError,
  isConsultChimpsError,
  OPERATION_ABORTED,
  type ByteArtifact,
  type ByteOperationOutcome,
  type OperationControlOptions,
} from "@consultchimps/core";

import { browserOutputPlace, removeOutputs } from "@/lib/output-storage";
import { unreadableFile } from "@/lib/unreadable-file";
import type {
  NamedFile,
  OperationTask,
  OutputFile,
  WorkerCommand,
  WorkerEvent,
} from "@/lib/operation-tasks";

/**
 * The worker global, typed locally. Pulling in the `webworker` lib would
 * collide with the DOM lib the rest of the app compiles against, and the two
 * members used here are the whole surface.
 */
const scope = self as unknown as {
  addEventListener(
    type: "message",
    listener: (event: MessageEvent<WorkerCommand>) => void,
  ): void;
  postMessage(message: WorkerEvent): void;
};

/**
 * The OPFS outputs this worker wrote, one set per file-producing run. A run's
 * outputs are deleted when the run after next starts, not the next one, so a
 * download of them still in progress is not cut short. A sweep removes any
 * that a closed tab left behind.
 */
const outputRuns: Set<string>[] = [];
const outputPlace = browserOutputPlace();
let sweeping = removeOutputs(outputPlace, new Set());

/** Start a run's set of outputs, deleting those two runs old. */
async function nextOutputRun(): Promise<Set<string>> {
  await sweeping;
  const stale = new Set(
    outputRuns
      .splice(0, Math.max(0, outputRuns.length - 1))
      .flatMap((run) => [...run]),
  );
  sweeping = removeOutputs(outputPlace, stale);
  await sweeping;
  // Names whose removal failed stay listed, so a later run retries them.
  if (stale.size > 0) outputRuns.unshift(stale);
  const run = new Set<string>();
  outputRuns.push(run);
  return run;
}

/** One controller per in-flight task, so a `cancel` command can reach it. */
const controllers = new Map<number, AbortController>();

interface TaskAnswer {
  readonly value: unknown;
  readonly artifacts?: readonly OutputFile[] | undefined;
}

/** A promise that rejects with the usual cancellation once `signal` aborts. */
function cancelled(signal: AbortSignal | undefined): Promise<never> {
  return new Promise((_, reject) => {
    if (signal === undefined) return;
    const fail = (): void => {
      reject(
        new ConsultChimpsError(
          OPERATION_ABORTED,
          "The task was cancelled before it finished. No source data was changed and no partial output was produced.",
        ),
      );
    };
    if (signal.aborted) fail();
    else signal.addEventListener("abort", fail, { once: true });
  });
}

/**
 * Read one input whole, for the operations that need every byte at once. The
 * copy lives here in the worker only; the page never held one.
 */
async function whole(
  input: NamedFile,
  signal: AbortSignal | undefined,
): Promise<{
  readonly name: string;
  readonly bytes: Uint8Array;
}> {
  // A task cancelled before it got here starts no read at all.
  if (signal?.aborted) return cancelled(signal);
  // A large or cloud-backed file can take a while; Cancel must not wait for it.
  const reading = input.file.arrayBuffer().then(
    (buffer) => ({ name: input.name, bytes: new Uint8Array(buffer) }),
    (error: unknown) => {
      throw unreadableFile(input.name, error);
    },
  );
  return Promise.race([reading, cancelled(signal)]);
}

/**
 * Hand an output over as a `Blob`. The page receives a reference to it, and
 * the buffer here is dropped once the task answers.
 */
function toOutputFile(artifact: ByteArtifact): OutputFile {
  return {
    name: artifact.name,
    blob: new Blob([artifact.bytes as Uint8Array<ArrayBuffer>], {
      type: artifact.mediaType ?? "",
    }),
    ...(artifact.mediaType === undefined
      ? {}
      : { mediaType: artifact.mediaType }),
  };
}

function answerWithOutputs<TMetric extends string>(
  outcome: ByteOperationOutcome<TMetric>,
): TaskAnswer {
  return {
    value: outcome.result,
    artifacts: outcome.outputs.map(toOutputFile),
  };
}

function answerWithValue(value: unknown): TaskAnswer {
  return { value };
}

async function perform(
  task: OperationTask,
  controls: Required<OperationControlOptions>,
): Promise<TaskAnswer> {
  switch (task.kind) {
    case "pdf.plan-split": {
      const { planSplitPdfBytes } = await import("@consultchimps/pdf/bytes");
      return answerWithValue(
        await planSplitPdfBytes({
          input: await whole(task.input, controls.signal),
          filenamePrefix: task.filenamePrefix,
        }),
      );
    }
    case "pdf.split": {
      const { splitPdfBytes } = await import("@consultchimps/pdf/bytes");
      return answerWithOutputs(
        await splitPdfBytes({
          ...controls,
          input: await whole(task.input, controls.signal),
          filenamePrefix: task.filenamePrefix,
        }),
      );
    }
    case "pdf.merge": {
      const { mergePdfsBytes } = await import("@consultchimps/pdf/bytes");
      return answerWithOutputs(
        await mergePdfsBytes({
          ...controls,
          inputs: await Promise.all(
            task.inputs.map((input) => whole(input, controls.signal)),
          ),
          outputName: task.outputName,
        }),
      );
    }
    case "xlsx.plan-split": {
      // Read in pieces through `Blob.slice`, never whole.
      const { planSplitFile } = await import("@/lib/streamed-split");
      return answerWithValue(
        await planSplitFile(task.input, task.options, controls.signal),
      );
    }
    case "xlsx.split": {
      // Read in pieces and each output written to disk as it goes.
      const { splitFile } = await import("@/lib/streamed-split");
      const { result, outputs } = await splitFile(
        task.input,
        task.options,
        controls,
        outputPlace,
        await nextOutputRun(),
      );
      return { value: result, artifacts: outputs };
    }
    case "xlsx.merge": {
      // Read in pieces and written to disk as it goes (ADR 0006).
      const { mergeFiles } = await import("@/lib/streamed-merge");
      const { result, outputs } = await mergeFiles(
        task.inputs,
        { csv: task.csv, outputName: task.outputName, values: task.values },
        controls,
        outputPlace,
        await nextOutputRun(),
      );
      return { value: result, artifacts: outputs };
    }
    case "xlsx.consolidate": {
      // Read in pieces and written to disk as it goes (ADR 0006).
      const { consolidateFiles } = await import("@/lib/streamed-consolidation");
      const { result, outputs } = await consolidateFiles(
        task.inputs,
        {
          addSourceColumns: task.addSourceColumns,
          csv: task.csv,
          includeHiddenSheets: task.includeHiddenSheets,
          mapping: task.mapping,
          normalizeHeaders: task.normalizeHeaders,
          outputName: task.outputName,
        },
        controls,
        outputPlace,
        await nextOutputRun(),
      );
      return { value: result, artifacts: outputs };
    }
    case "xlsx.suggest-mapping": {
      // The suggestion is drafted from the tables the consolidation read, so
      // the page proposes exactly what the library proposes for these
      // workbooks and these options, from the first pass alone: nothing is
      // written.
      const { suggestMappingFromFiles } =
        await import("@/lib/streamed-consolidation");
      return answerWithValue(
        await suggestMappingFromFiles(
          task.inputs,
          task.includeHiddenSheets,
          controls,
          task.csv,
        ),
      );
    }
    case "xlsx.columns": {
      if (task.input.name.toLowerCase().endsWith(".csv")) {
        // A CSV file's one worksheet, read in pieces; its header row's
        // columns are what the inspection reports.
        const { inspectFile } = await import("@/lib/streamed-inspection");
        const outcome = await inspectFile(task.input, {
          ...controls,
          csv: task.csv,
          headerRow: task.headerRow,
          sampleValues: 0,
        });
        const sheet = outcome.description.sheets[0];
        return answerWithValue({
          columns: sheet?.columns.map((column) => column.header) ?? [],
          worksheet: sheet?.name ?? "",
        });
      }
      const { readWorksheetRecordsBytes } =
        await import("@consultchimps/xlsx/bytes");
      const records = await readWorksheetRecordsBytes(
        await whole(task.input, controls.signal),
        {
          headerRow: task.headerRow,
          worksheet: task.worksheet,
        },
      );
      return answerWithValue({
        columns: records.columns,
        worksheet: records.worksheet,
      });
    }
    case "xlsx.inspect": {
      // Read in pieces through `Blob.slice`, never whole.
      const { inspectFile } = await import("@/lib/streamed-inspection");
      return answerWithValue(
        await inspectFile(task.input, { ...controls, ...task.options }),
      );
    }
    case "xlsx.unprotect": {
      const { unprotectWorkbookBytes } =
        await import("@consultchimps/xlsx/bytes");
      return answerWithOutputs(
        await unprotectWorkbookBytes({
          ...controls,
          input: await whole(task.input, controls.signal),
          ...(task.outputName === undefined
            ? {}
            : { outputName: task.outputName }),
        }),
      );
    }
    case "pptx.inspect": {
      const { inspectPresentationOutcomeBytes } =
        await import("@consultchimps/pptx/bytes");
      return answerWithValue(
        await inspectPresentationOutcomeBytes(
          await whole(task.template, controls.signal),
          {
            ...controls,
            templateSlide: task.templateSlide,
          },
        ),
      );
    }
    case "pptx.plan-populate": {
      const { planPopulatePresentationBytes } =
        await import("@consultchimps/pptx/bytes");
      return answerWithValue(
        await planPopulatePresentationBytes({
          ...controls,
          ...task.options,
          template: await whole(task.template, controls.signal),
          workbook: await whole(task.workbook, controls.signal),
        }),
      );
    }
    case "pptx.populate": {
      const { populatePresentationBytes } =
        await import("@consultchimps/pptx/bytes");
      return answerWithOutputs(
        await populatePresentationBytes({
          ...controls,
          ...task.options,
          template: await whole(task.template, controls.signal),
          workbook: await whole(task.workbook, controls.signal),
        }),
      );
    }
  }
  throw new Error("Unsupported operation task");
}

async function execute(id: number, task: OperationTask): Promise<void> {
  const controller = new AbortController();
  controllers.set(id, controller);
  try {
    const answer = await perform(task, {
      onProgress: (progress) => {
        scope.postMessage({ type: "progress", id, progress });
      },
      signal: controller.signal,
    });
    scope.postMessage({
      type: "done",
      id,
      value: answer.value,
      artifacts: answer.artifacts,
    });
  } catch (error) {
    scope.postMessage({
      type: "failed",
      id,
      message:
        error instanceof Error
          ? error.message
          : "An unexpected problem occurred.",
      code: isConsultChimpsError(error) ? error.code : undefined,
    });
  } finally {
    controllers.delete(id);
  }
}

scope.addEventListener("message", (event) => {
  const command = event.data;
  if (command.type === "release") {
    // The page no longer offers them: delete them, as far as time allows.
    // Names whose removal failed go back on the list for the next try.
    const offered = new Set(outputRuns.splice(0).flatMap((run) => [...run]));
    sweeping = sweeping
      .then(() => removeOutputs(outputPlace, offered))
      .then(() => {
        if (offered.size > 0) outputRuns.unshift(offered);
      });
    return;
  }
  if (command.type === "cancel") {
    controllers.get(command.id)?.abort();
    return;
  }
  void execute(command.id, command.task);
});
